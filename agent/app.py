from __future__ import annotations

import io
import hashlib
import json
import os
import re
import shutil
import tempfile
import time
import uuid
import zipfile
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from pathlib import Path, PurePosixPath
from typing import Any

from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request, send_file
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from docx import Document
from docx.shared import Inches, Pt

from .deepseek_vision import DeepSeekError, DeepSeekVision

ROOT = Path(__file__).resolve().parents[1]
# This local app is configured through its project .env file. Prefer that
# explicit configuration so stale user-level credentials cannot mask it.
load_dotenv(ROOT / ".env", override=True)
DATA_DIR = ROOT / "data"
ROSTER_PATH = DATA_DIR / "operator_roster.json"
DEVICE_IDS = [f"DCM-{n}" for n in range(1, 9)]
FIELDS = ["Date", "Pile NO.", "Equipment NO.", "Drilling Start Time",
          "Drilling Complete Time", "Point Complete Time", "Operator", "Cement Content"]
MAX_UPLOAD = int(os.getenv("MAX_UPLOAD_MB", "100")) * 1024 * 1024
VISION_CONFIDENCE_THRESHOLD = float(os.getenv("VISION_CONFIDENCE_THRESHOLD", "0.75"))


def normalize_pile_no(value: Any) -> str:
    """Pad a final numeric pile segment to four digits without changing its prefix."""
    text = str(value or "").strip()
    match = re.fullmatch(r"(.*-)(\d+)", text)
    if not match:
        return text
    prefix, number = match.groups()
    if len(number) < 4:
        number = number.zfill(4)
    return prefix + number

app = Flask(__name__, template_folder=str(ROOT / "templates"), static_folder=str(ROOT / "static"))
app.secret_key = os.getenv("APP_SECRET_KEY", "local-agent-change-me")
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD

if not ROSTER_PATH.exists():
    DATA_DIR.mkdir(exist_ok=True)
    ROSTER_PATH.write_text(json.dumps({"roster":[
        {"account":"Raj", "full_name":"Raj Pillai"},
        {"account":"Dev", "full_name":"Dev Sharma"},
        {"account":"Arjun", "full_name":"Arjun Mehra"},
        {"account":"Vikram", "full_name":"Vikram Singh"},
        {"account":"Aditya", "full_name":"Aditya Sharma"},
        {"account":"Rohit", "full_name":"Rohit Verma"},
        {"account":"Karan", "full_name":"Karan Malhotra"},
        {"account":"Sidd", "full_name":"Siddharth Krishnamurthy"}
    ]}, ensure_ascii=False, indent=2), encoding="utf-8")

# The prototype is local and single-user. Stage state is isolated by run_id.
RUNS: dict[str, dict[str, Any]] = {}
JOBS: dict[str, dict[str, Any]] = {}
JOB_LOCK = threading.Lock()
IMAGE_POOL = ThreadPoolExecutor(max_workers=2, thread_name_prefix="image-vision")


def new_run(report_date: str) -> dict[str, Any]:
    datetime.strptime(report_date, "%Y-%m-%d")
    run = {"run_id":uuid.uuid4().hex, "report_date":report_date,
           "construction":{}, "operation":[], "silo":[], "manual":[], "raw":{}, "artifacts":{}, "warnings":[]}
    RUNS[run["run_id"]] = run
    return run


def get_run(run_id: str) -> dict[str, Any]:
    if not run_id or run_id not in RUNS:
        raise ValueError("处理会话已过期，请重新从日报日期开始。")
    return RUNS[run_id]


def parse_chat(raw: str, report_date: str | None = None) -> list[dict[str, Any]]:
    """Parse common Android/iOS WhatsApp exports, including US M/D/YY AM/PM."""
    invisible = str.maketrans("", "", "\u200e\u200f\ufeff")
    header = re.compile(r"^\[?(\d{1,4}[/-]\d{1,2}[/-]\d{2,4}),?\s+(\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?)\]?\s*(?:(?:-|–)\s*|\s+)([^:]{1,100}):\s?(.*)$")
    out: list[dict[str, Any]] = []
    cur: dict[str, Any] | None = None
    def stamp(ds: str, ts: str) -> str | None:
        ds = ds.strip(); ts = ts.strip().upper().replace(" ", " ")
        formats = ["%d/%m/%Y", "%d/%m/%y", "%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%d-%m-%Y", "%m-%d-%Y"]
        parsed_dates = []
        for fmt in formats:
            try:
                candidate=datetime.strptime(ds, fmt).date()
                if candidate not in parsed_dates: parsed_dates.append(candidate)
            except ValueError: pass
        parsed_date = None
        if report_date and parsed_dates:
            target=date.fromisoformat(report_date)
            allowed={target,target+timedelta(days=1)}
            parsed_date=next((candidate for candidate in parsed_dates if candidate in allowed),None)
        if parsed_date is None and parsed_dates: parsed_date=parsed_dates[0]
        if not parsed_date: return None
        for fmt in ("%H:%M:%S", "%H:%M", "%I:%M:%S %p", "%I:%M %p"):
            try: return datetime.combine(parsed_date, datetime.strptime(ts, fmt).time()).isoformat()
            except ValueError: pass
        return None
    for line in raw.translate(invisible).splitlines():
        match = header.match(line)
        if match:
            if cur: out.append(cur)
            ds, ts, sender, body = match.groups()
            cur = {"message_time":stamp(ds, ts), "sender":sender.strip(), "text":body.strip(), "media_ref":None}
            attach = re.search(r"<attached:\s*([^>]+)>", body, re.I)
            if attach: cur["media_ref"] = attach.group(1).strip()
            else:
                filename=re.search(r"([\w(). -]+\.(?:jpe?g|png|webp|bmp))",body,re.I)
                if filename: cur["media_ref"]=filename.group(1).strip()
        elif cur and line.strip():
            cur["text"] = (cur.get("text", "") + "\n" + line.strip()).strip()
            attach = re.search(r"<attached:\s*([^>]+)>", line, re.I)
            if attach: cur["media_ref"] = attach.group(1).strip()
            else:
                filename=re.search(r"([\w(). -]+\.(?:jpe?g|png|webp|bmp))",line,re.I)
                if filename: cur["media_ref"]=filename.group(1).strip()
    if cur: out.append(cur)
    return out


def in_window(messages: list[dict[str, Any]], report_date: str) -> list[dict[str, Any]]:
    start = datetime.combine(date.fromisoformat(report_date), datetime.min.time()).replace(hour=7)
    end = start + timedelta(days=1)
    return [m for m in messages if not m.get("message_time") or start <= datetime.fromisoformat(m["message_time"]) < end]


def read_upload_text(file) -> str:
    data = file.read()
    try: return data.decode("utf-8-sig")
    except UnicodeDecodeError: return data.decode("utf-16", errors="replace")


def parse_clock(raw: str | None, report_date: str) -> tuple[str | None, str | None]:
    if not raw: return None, None
    value = raw.strip().lower().replace("hrs", "").replace("hr", "").replace("hours", "").strip()
    match = re.fullmatch(r"(\d{1,2}):?(\d{2})(?::(\d{2}))?", value)
    if not match: return None, None
    hour, minute = int(match.group(1)), int(match.group(2))
    if hour > 23 or minute > 59: return None, None
    clock = f"{hour:02d}:{minute:02d}"
    anchor = date.fromisoformat(report_date)
    dt = datetime.combine(anchor, datetime.strptime(clock, "%H:%M").time())
    if hour < 7: dt += timedelta(days=1)
    return clock, dt.isoformat(timespec="minutes")


def parse_construction_text(raw: str, device: str, report_date: str) -> list[dict[str, Any]]:
    messages = parse_chat(raw, report_date)
    if not messages:
        messages = [{"message_time":None,"sender":None,"text":raw,"media_ref":None}]
    messages = in_window(messages, report_date)
    fields = {
        "pile":"Point\\s*No\\s*:", "date":"Date\\s*:",
        "start":"Drilling\\s+start\\s+time\\s*:",
        "drill":"Drilling\\s+complete\\s+time\\s*:",
        "complete":"Point\\s+complete\\s+time\\s*:",
    }
    result=[]
    for message in messages:
        text=message.get("text") or ""
        def field(name):
            m=re.search(fields[name]+r"\s*([^\r\n]+)",text,re.I)
            return m.group(1).strip().rstrip(" .") if m else None
        raw_pile=field("pile")
        pile=normalize_pile_no(raw_pile)
        if not pile: continue
        st,st_dt=parse_clock(field("start"),report_date)
        dr,dr_dt=parse_clock(field("drill"),report_date)
        co,co_dt=parse_clock(field("complete"),report_date)
        result.append({"Date":report_date,"Pile NO.":pile,"Equipment NO.":device,
          "Drilling Start Time":st,"Drilling Complete Time":dr,"Point Complete Time":co,
          "Operator":"","Cement Content":"","_start_dt":st_dt,"_drill_dt":dr_dt,"_complete_dt":co_dt,
          "_message_time":message.get("message_time"),"_source_text":text,"_source_group":device,
          "_pile_source_raw":raw_pile,"_pile_normalized":pile!=raw_pile})
    return result


def _zip_assets(file) -> tuple[str, list[str], list[str]]:
    temp=tempfile.mkdtemp(prefix="site_agent_")
    txt_paths=[]; images=[]
    try:
        with zipfile.ZipFile(io.BytesIO(file.read())) as zf:
            for info in zf.infolist():
                name=info.filename.replace("\\", "/")
                path=PurePosixPath(name)
                if path.is_absolute() or ".." in path.parts or info.is_dir(): continue
                target=Path(temp).joinpath(*path.parts)
                target.parent.mkdir(parents=True,exist_ok=True)
                with zf.open(info) as src, target.open("wb") as dst: shutil.copyfileobj(src,dst)
                if target.suffix.lower()==".txt": txt_paths.append(str(target))
                elif target.suffix.lower() in (".jpg",".jpeg",".png",".webp",".bmp"): images.append(str(target))
    except (zipfile.BadZipFile, OSError):
        shutil.rmtree(temp,ignore_errors=True)
        raise ValueError("文件不是有效 ZIP，或 ZIP 内容无法读取。")
    return temp,txt_paths,images


def _chat_from_zip(txt_paths: list[str], report_date: str) -> tuple[str,list[dict[str,Any]]]:
    if not txt_paths: raise ValueError("ZIP 中没有找到 WhatsApp 导出的 .txt 聊天记录。")
    # Prefer the likely chat file; combine only if exports were split across files.
    raw="\n".join(Path(p).read_text(encoding="utf-8-sig",errors="replace") for p in txt_paths)
    return raw,parse_chat(raw,report_date)


def _media_for_message(message: dict[str,Any], images: list[str], used: set[str]) -> str | None:
    ref=message.get("media_ref")
    if ref:
        for path in images:
            if Path(path).name.lower()==Path(ref).name.lower(): used.add(path); return path
    for path in images:
        if path not in used: used.add(path); return path
    return None


def _vision(stage: str, prompt: str, image_path: str) -> dict[str,Any]:
    return DeepSeekVision().extract_stage_json(stage,prompt,image_path)


def load_roster() -> list[dict[str,str]]:
    try: return json.loads(ROSTER_PATH.read_text(encoding="utf-8")).get("roster",[])
    except Exception: return []


def _feishu_context():
    import requests
    app_id=os.getenv("LARK_APP_ID"); secret=os.getenv("LARK_APP_SECRET")
    app_token=os.getenv("LARK_BASE_APP_TOKEN"); table_id=os.getenv("LARK_TABLE_ID")
    if not all((app_id,secret,app_token,table_id)):
        raise ValueError("飞书配置未齐全：需要 App ID、App Secret、Bitable App Token 与 Table ID。")
    domain=os.getenv("LARK_DOMAIN","https://open.feishu.cn").rstrip("/")
    res=requests.post(domain+"/open-apis/auth/v3/tenant_access_token/internal",json={"app_id":app_id,"app_secret":secret},timeout=20)
    res.raise_for_status(); data=res.json()
    if data.get("code")!=0: raise RuntimeError(data.get("msg","获取飞书 token 失败"))
    return requests,domain,data["tenant_access_token"],app_token,table_id


def _operator_field(requests,domain,token,app_token,table_id):
    url=f"{domain}/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/fields"
    res=requests.get(url,headers={"Authorization":f"Bearer {token}"},params={"page_size":100},timeout=20)
    res.raise_for_status(); data=res.json()
    if data.get("code")!=0: raise RuntimeError(data.get("msg","读取飞书字段失败"))
    return next((field for field in data.get("data",{}).get("items",[]) if field.get("field_name")=="Operator"),None)


def _add_operator_options(options:list[str]):
    requests,domain,token,app_token,table_id=_feishu_context()
    field=_operator_field(requests,domain,token,app_token,table_id)
    if not field: raise RuntimeError("飞书表中找不到名为 Operator 的字段。")
    if field.get("type")!=3 and field.get("ui_type")!="SingleSelect":
        raise RuntimeError("飞书 Operator 字段不是单选字段，未更改字段配置。")
    prop=dict(field.get("property") or {}); current=prop.get("options",[])
    existing={str(x.get("name","")).casefold() for x in current}
    added=[name for name in options if name and name.casefold() not in existing]
    if not added: return []
    prop["options"]=[*current,*[{"name":name,"color":0} for name in added]]
    url=f"{domain}/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/fields/{field['field_id']}"
    res=requests.put(url,headers={"Authorization":f"Bearer {token}"},json={"field_name":"Operator","type":field.get("type",3),"property":prop},timeout=20)
    res.raise_for_status(); data=res.json()
    if data.get("code")!=0: raise RuntimeError(data.get("msg","新增 Operator 选项失败"))
    return added


def match_operator(account: str | None) -> str:
    if not account: return ""
    key=account.strip().casefold()
    for item in load_roster():
        if item.get("account","").strip().casefold()==key: return item.get("full_name","")
    # Treat a sender name that already equals a roster full name as exact.
    for item in load_roster():
        if item.get("full_name","").strip().casefold()==key: return item.get("full_name","")
    return ""


def safe_rows(run: dict[str,Any]) -> list[dict[str,Any]]:
    merged: dict[str,dict[str,Any]]={}
    for source in [*run["construction"].values(),run["operation"],run["silo"],run["manual"]]:
        for row in source:
            pile=normalize_pile_no(row.get("Pile NO."))
            if not pile: continue
            if pile not in merged:
                merged[pile]={k:row.get(k,"") for k in FIELDS}
                merged[pile]["Pile NO."]=pile
                merged[pile].update({k:v for k,v in row.items() if k.startswith("_")})
            else:
                target=merged[pile]
                for key in FIELDS:
                    if row.get(key) not in (None,""):
                        if target.get(key) in (None,""): target[key]=row[key]
                for key,value in row.items():
                    if key.startswith("_") and not target.get(key): target[key]=value
    return sorted(merged.values(),key=lambda r:(r.get("_complete_dt") or "9999",r.get("Pile NO.") or ""))


def build_warnings(run: dict[str,Any]) -> list[dict[str,str]]:
    warnings=[]
    for group in run["construction"].values():
        for record in group:
            missing=[field for field in ("Drilling Start Time","Drilling Complete Time","Point Complete Time") if not record.get(field)]
            if missing: warnings.append({"type":"missing","title":"施工时间信息缺失","detail":f"{record.get('Pile NO.')} 缺少 {', '.join(missing)}。"})
            dts=[record.get("_start_dt"),record.get("_drill_dt"),record.get("_complete_dt")]
            try:
                parsed=[datetime.fromisoformat(x) for x in dts if x]
                if len(parsed)==3 and not (parsed[0]<=parsed[1]<=parsed[2]):
                    warnings.append({"type":"time","title":"施工时间顺序异常","detail":f"{record.get('Pile NO.')} 的开始、钻进完成和桩点完成时间顺序需要核对。"})
            except ValueError: pass
    by_pile:dict[str,list[dict[str,Any]]]={}
    for source in [*run["construction"].values(),run["operation"],run["silo"]]:
        for row in source:
            if row.get("Pile NO."): by_pile.setdefault(str(row["Pile NO."]),[]).append(row)
    for pile,items in by_pile.items():
        for field in ("Operator","Cement Content"):
            values={str(r.get(field)) for r in items if r.get(field) not in (None,"")}
            if len(values)>1: warnings.append({"type":"conflict","title":"桩号信息冲突","detail":f"{pile} 对应的 {field} 出现不同内容，请工程师核对。"})
        operators={str(r.get("_account") or r.get("Operator")) for r in items if r.get("_account") or r.get("Operator")}
        if len(operators)>1: warnings.append({"type":"conflict","title":"操作员信息冲突","detail":f"{pile} 对应多个操作员账户。"})
        images={r.get("_image_hash") for r in items if r.get("_image_hash")}
        if len(images)>1: warnings.append({"type":"conflict","title":"图片信息冲突","detail":f"{pile} 对应不同图片，请工程师核对。"})
    for row in run["operation"]:
        try: confidence=float(row.get("_vision_confidence"))
        except (TypeError,ValueError): confidence=0.0
        if confidence<VISION_CONFIDENCE_THRESHOLD:
            warnings.append({"type":"missing","title":"图片识别置信度较低","detail":f"Operation 图片对应桩号 {row.get('Pile NO.') or '未识别'} 的置信度为 {confidence:.2f}，请在审核 Word 中对照原图。"})
        mt=row.get("_message_time"); pt=next((item.get("_complete_dt") for group in run["construction"].values() for item in group if item.get("Pile NO.")==row.get("Pile NO.")),None)
        if mt and pt:
            try:
                mins=abs((datetime.fromisoformat(mt)-datetime.fromisoformat(pt)).total_seconds())/60
                if mins>20: warnings.append({"type":"time","title":"上传时间需要核对","detail":f"{row.get('Pile NO.')} 的图片发送时间与施工完成时间相差超过 20 分钟。"})
            except ValueError: pass
    if run["artifacts"].get("operation") and not run["operation"]:
        warnings.append({"type":"missing","title":"未识别到目标设备图片","detail":"Operation ZIP 中没有识别到可用的机器屏幕桩号记录。请核对图片和聊天导出内容。"})
    if run["operation"] and run["construction"]:
        construction_piles={r.get("Pile NO.") for grp in run["construction"].values() for r in grp}
        for row in run["operation"]:
            if row.get("Pile NO.") and row["Pile NO."] not in construction_piles:
                warnings.append({"type":"missing","title":"桩号未匹配施工记录","detail":f"{row['Pile NO.']} 尚未在 8 个施工群聊中找到对应记录。"})
            if not row.get("Operator"):
                warnings.append({"type":"missing","title":"操作员全名未匹配","detail":f"账户“{row.get('_account','')}”尚未匹配到操作员映射表中的全名。"})
    if run["artifacts"].get("silo") and not run["silo"]:
        warnings.append({"type":"missing","title":"未识别到水泥仪表记录","detail":"Silo ZIP 中没有识别到可用的仪表读数和对应桩号。请核对图片与消息。"})
    if run["silo"]:
        construction_piles={r.get("Pile NO.") for group in run["construction"].values() for r in group}
        for row in run["silo"]:
            try: confidence=float(row.get("_vision_confidence"))
            except (TypeError,ValueError): confidence=0.0
            if confidence<VISION_CONFIDENCE_THRESHOLD:
                warnings.append({"type":"missing","title":"图片识别置信度较低","detail":f"Silo 图片对应桩号 {row.get('Pile NO.') or '未识别'} 的置信度为 {confidence:.2f}，请在审核 Word 中对照原图。"})
            if row.get("Pile NO.") and row.get("Pile NO.") not in construction_piles:
                warnings.append({"type":"missing","title":"水泥记录未匹配施工桩号","detail":f"{row['Pile NO.']} 尚未找到施工记录。"})
    run["warnings"]=warnings
    return warnings


def analysis_for(run: dict[str,Any]) -> list[dict[str,Any]]:
    grouped:dict[tuple[str,str],list[dict[str,Any]]]={}
    for row in safe_rows(run):
        name=row.get("Operator")
        if not name: continue
        start=row.get("_start_dt"); end=row.get("_complete_dt")
        try: hours=max(0,(datetime.fromisoformat(end)-datetime.fromisoformat(start)).total_seconds()/3600) if start and end else 0
        except ValueError: hours=0
        h=int(start[11:13]) if start and len(start)>12 else 0
        shift="白班" if 7<=h<19 else "夜班"
        grouped.setdefault((name,shift),[]).append({"hours":hours,"pile":row.get("Pile NO.")})
    return [{"operator":name,"shift":shift,"piles":len(rows),"hours":round(sum(x["hours"] for x in rows),2),"avg":round(sum(x["hours"] for x in rows)/len(rows),2)} for (name,shift),rows in grouped.items()]


def _write_xlsx(rows: list[dict[str,Any]]) -> io.BytesIO:
    wb=Workbook(); ws=wb.active; ws.title="Construction Review"
    ws.append(FIELDS[:6])
    for row in rows: ws.append([row.get(k) or "" for k in FIELDS[:6]])
    for c in ws[1]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="2457D6"); c.alignment=Alignment(horizontal="center")
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    widths=[14,18,16,24,28,24]
    for i,w in enumerate(widths,1): ws.column_dimensions[chr(64+i)].width=w
    for row in ws.iter_rows(min_row=2):
        for cell in row: cell.alignment=Alignment(vertical="center")
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf


def _write_txt_zip(run: dict[str,Any]) -> io.BytesIO:
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as zf:
        active_devices=[device for device in DEVICE_IDS if device in run["construction"]]
        inactive_devices=[device for device in DEVICE_IDS if device not in run["construction"]]
        all_records=[row for device in DEVICE_IDS for row in run["construction"].get(device,[])]
        all_records.sort(key=lambda row:(row.get("_complete_dt") or "9999",row.get("Pile NO.") or ""))
        summary=["施工审核材料 · 当天工作设备汇总",f"日报日期：{run['report_date']}",f"有效时间窗：{run['report_date']} 07:00 至次日 07:00",f"当天工作设备（已上传）：{'、'.join(active_devices)}",f"当天未工作设备（未上传）：{'、'.join(inactive_devices) if inactive_devices else '无'}",f"提取桩点总数：{len(all_records)}","排序依据：Point Complete Time（24 小时制）","","序号 | Point Complete Time | Pile NO. | Equipment NO. | Drilling Start | Drilling Complete","-"*102]
        for number,row in enumerate(all_records,1):
            summary.append(f"{number:02d} | {row.get('Point Complete Time') or '未识别':>8} | {row.get('Pile NO.') or '未识别'} | {row.get('Equipment NO.') or '未识别'} | {row.get('Drilling Start Time') or '未识别'} | {row.get('Drilling Complete Time') or '未识别'}")
        summary.extend(["","核对说明：请先用本汇总确认桩点数量和完成顺序，再打开各设备材料核对完整施工信息。"])
        zf.writestr("00_All_Equipment_Summary.txt","\ufeff"+"\n".join(summary))
        for device in active_devices:
            records=run["construction"].get(device,[])
            lines=[f"施工审核证据材料 · {device}",f"日报日期：{run['report_date']}",f"有效时间窗：{run['report_date']} 07:00 至次日 07:00",f"该设备有效桩点：{len(records)} 条","时间格式：24 小时制","="*72,"" ]
            for number,row in enumerate(records,1):
                lines.extend([f"记录 {number}｜{row.get('Pile NO.') or '未识别桩号'}","-"*72,"【写入预览表的关键信息】",f"Date                  : {row.get('Date','')}",f"Pile NO.              : {row.get('Pile NO.','')}",
                    f"Equipment NO.         : {row.get('Equipment NO.','')}",f"Drilling Start Time   : {row.get('Drilling Start Time') or '未识别'}",
                    f"Drilling Complete Time: {row.get('Drilling Complete Time') or '未识别'}",f"Point Complete Time   : {row.get('Point Complete Time') or '未识别'}","",
                    "【辅助判断信息｜有效消息正文，已移除 WhatsApp 聊天标题】",row.get("_source_text") or "（无消息正文）","","="*72,"" ])
            if not records: lines.append("该设备没有提取到位于所选日报时间窗内的有效桩点记录。")
            zf.writestr(f"{device}_Review.txt","\ufeff"+"\n".join(lines))
    buf.seek(0); return buf


def _write_review_docx(records: list[dict[str,Any]], heading: str, report_date: str, include_operation_details: bool = True) -> io.BytesIO:
    doc=Document(); doc.add_heading(heading,0)
    doc.add_paragraph(f"日报日期：{report_date}（07:00 至次日 07:00）")
    doc.add_paragraph(f"目标记录：{len(records)} 条。请将下列证据与网页飞书格式预览逐条对照；本文件不代表系统自动审核结论。")
    doc.add_paragraph("说明：文档只整理与目标记录相关的有效信息，原图保持不变。")
    for n,row in enumerate(records,1):
        if n>1: doc.add_page_break()
        doc.add_heading(f"记录 {n} · {row.get('Pile NO.') or '待识别桩号'}",2)
        table=doc.add_table(rows=0,cols=2); table.style="Table Grid"
        def add(label: str,value: Any):
            cells=table.add_row().cells; cells[0].text=label; cells[1].text=str(value if value not in (None,"") else "未识别")
        if include_operation_details:
            add("WhatsApp 账户名称",row.get("_account")); add("映射后的操作员全名",row.get("Operator")); add("图片发送时间",(row.get("_message_time") or "").replace("T"," ")); add("AI 提取桩号",row.get("Pile NO.")); add("识别说明",row.get("_recognized"))
            if row.get("_source_text"): add("图片所附消息",row.get("_source_text"))
        else:
            add("聊天文字中的桩号",row.get("Pile NO.")); add("仪表原始读数",row.get("_meter_reading")); add("写入 Cement Content（原始读数 ÷ 2）",row.get("Cement Content")); add("识别说明",row.get("_recognized")); add("与图片关联的消息文字",row.get("_source_text"))
        doc.add_heading("原始证据图片",3)
        if row.get("_image_path") and os.path.exists(row["_image_path"]):
            try: doc.add_picture(row["_image_path"],width=Inches(5.6))
            except Exception: doc.add_paragraph("（原图无法嵌入文档）")
    buf=io.BytesIO(); doc.save(buf); buf.seek(0); return buf


@app.get("/")
def home(): return render_template("index.html")


@app.get("/api/status")
def status():
    vision=DeepSeekVision()
    return jsonify({"model_provider":"DeepSeek","model":vision.model,"model_key":vision.available(),"feishu":bool(os.getenv("LARK_APP_ID") and os.getenv("LARK_APP_SECRET")),"ready":vision.available()})


@app.post("/api/runs")
def create_run():
    body=request.get_json(silent=True) or {}
    try: run=new_run(body.get("report_date",""))
    except Exception: return jsonify({"error":"请选择有效的日报日期。"}),400
    return jsonify({"run_id":run["run_id"],"report_date":run["report_date"],"window":f"{run['report_date']} 07:00 – 次日 07:00"})


@app.post("/api/runs/<run_id>/construction/<device>")
def upload_construction(run_id,device):
    try:
        run=get_run(run_id)
        if device not in DEVICE_IDS: return jsonify({"error":"设备编号无效。"}),400
        file=request.files.get("file")
        if not file or not file.filename.lower().endswith(".txt"): return jsonify({"error":"请选择 WhatsApp 导出的 .txt 文件。"}),400
        raw=read_upload_text(file); run["raw"][device]=raw
        rows=parse_construction_text(raw,device,run["report_date"])
        run["construction"][device]=rows
        all_rows=safe_rows(run); warnings=build_warnings(run)
        return jsonify({"device":device,"count":len(rows),"rows":all_rows,"warnings":warnings,"analysis":analysis_for(run),"message":f"{device} 解析完成，提取 {len(rows)} 条施工记录。"})
    except ValueError as e: return jsonify({"error":str(e)}),400


@app.post("/api/runs/<run_id>/construction/batch")
def upload_construction_batch(run_id):
    """Validate the files selected for working devices before committing them."""
    try:
        run=get_run(run_id); pending_raw={}; pending_rows={}
        selected=[device for device in DEVICE_IDS if request.files.get(device)]
        if not selected: return jsonify({"error":"请至少选择一台当天工作的设备 TXT 文件。"}),400
        for device in selected:
            file=request.files.get(device)
            if not file.filename.lower().endswith(".txt"): return jsonify({"error":f"{device} 不是 TXT 文件。"}),400
            raw=read_upload_text(file); parsed=parse_construction_text(raw,device,run["report_date"])
            if not parsed: return jsonify({"error":f"{device} 在所选日报时间窗内没有提取到有效桩点，请核对文件和日期。"}),400
            pending_raw[device]=raw; pending_rows[device]=parsed
        run["raw"].update(pending_raw); run["construction"]=pending_rows; run["artifacts"]["construction"]=True
        all_rows=safe_rows(run); warning_list=build_warnings(run)
        return jsonify({"device_count":len(selected),"count":sum(len(x) for x in pending_rows.values()),"rows":all_rows,"warnings":warning_list,"analysis":analysis_for(run),"message":f"{len(selected)} 台工作设备文件已确认并处理完成，共提取 {sum(len(x) for x in pending_rows.values())} 条桩点记录。"})
    except ValueError as e: return jsonify({"error":str(e)}),400


def _extract_media_zip(run:dict[str,Any],file,stage:str,on_progress=None):
    temp,txt_paths,images=_zip_assets(file)
    try:
        raw,messages=_chat_from_zip(txt_paths,run["report_date"])
        messages=in_window(parse_chat(raw,run["report_date"]),run["report_date"])
        if not messages: raise ValueError("选择的日报时间范围内没有可用 WhatsApp 消息。请检查日期与 ZIP 内容。")
        used:set[str]=set(); tasks=[]
        operation_prompt=("你在审核 WhatsApp Operation 群导出的原始施工设备照片。判断图片是否显示 D.S.M PLANT OVERVIEW 控制屏，并读取 Cluster No. Input / No. 输入框中的完整桩号。图片可能是后端自动裁剪出的右上区域，也可能是原图。不要把时间、设备标签或其他数字当作桩号。忽略纸张、仪表表盘和普通现场照。只返回 JSON：{\"is_target\":true/false,\"pile_no\":\"识别出的完整桩号或空字符串\",\"confidence\":0到1,\"note\":\"简短原因\"}。逐字符核对字母、数字和连字符；不确定时留空，禁止猜测。")
        silo_prompt=("你在审核 WhatsApp Silo 水泥群导出的原始设备照片。只接受方形或矩形的电子水泥计数器/控制仪表，读取红色 LED 当前真正点亮的完整数字。七段数码管未点亮的段可能留下暗红色 8 字轮廓，必须完全忽略；例如前两位只有暗淡的 88 轮廓、后两位明亮显示 33 时，读数是 33，不是 8833 或 8888。图片可能是后端自动生成的高亮段掩膜，也可能是原图。忽略机械水表、红色指针、时间戳、油漆标记、普通现场照和文字标签。只返回 JSON：{\"is_target\":true/false,\"red_number\":数字或null,\"confidence\":0到1,\"note\":\"简短原因\"}。逐位核对；无法确认时返回 null，禁止猜测。")
        for i,msg in enumerate(messages):
            body=msg.get("text") or ""
            if not msg.get("media_ref") and not re.search(r"<media omitted>|image omitted|photo omitted|\bimage\b|\bphoto\b",body,re.I): continue
            path=_media_for_message(msg,images,used)
            if not path: continue
            task={"index":i,"message":msg,"path":path,"pile_source":"","raw_pile":""}
            if stage=="silo":
                pile_source=body
                pile_match=re.search(r"(?:pile\s*(?:no\.?|id)?|桩号)\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9\- ]{2,})",pile_source,re.I)
                if not pile_match:
                    for prior in reversed(messages[max(0,i-4):i]):
                        if prior.get("sender")==msg.get("sender"):
                            pile_match=re.search(r"(?:pile\s*(?:no\.?|id)?|桩号)\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9\- ]{2,})",prior.get("text") or "",re.I)
                            if pile_match: pile_source=prior.get("text") or ""; break
                if not pile_match: continue
                task["pile_source"]=pile_source
                task["raw_pile"]=pile_match.group(1).strip().rstrip(".,;")
                task["prompt"]=silo_prompt
            else: task["prompt"]=operation_prompt
            tasks.append(task)

        total=len(tasks)
        if on_progress: on_progress(0,total,"准备识别图片")
        results={}; errors=[]; completed=0
        futures={IMAGE_POOL.submit(_vision,stage,task["prompt"],task["path"]):task for task in tasks}
        for future in as_completed(futures):
            task=futures[future]; completed+=1
            try: results[task["index"]]=future.result()
            except Exception as exc: errors.append(f"{Path(task['path']).name}：{exc}")
            if on_progress: on_progress(completed,total,Path(task["path"]).name)

        records=[]
        for task in tasks:
            vision=results.get(task["index"])
            if not vision or not vision.get("is_target"): continue
            msg=task["message"]; path=task["path"]
            fingerprint=hashlib.sha256(Path(path).read_bytes()).hexdigest()
            if stage=="operation":
                account=msg.get("sender") or ""
                raw_pile=str(vision.get("pile_no") or "").strip(); pile=normalize_pile_no(raw_pile)
                full=match_operator(account)
                normalization_note=f"（标准化自 {raw_pile}）" if pile and pile!=raw_pile else ""
                records.append({"Date":run["report_date"],"Pile NO.":pile,"Equipment NO.":"","Drilling Start Time":"","Drilling Complete Time":"","Point Complete Time":"","Operator":full,"Cement Content":"","_account":account,"_message_time":msg.get("message_time"),"_image_path":path,"_image_hash":fingerprint,"_vision_confidence":vision.get("confidence"),"_vision_strategy":vision.get("_vision_strategy","original"),"_recognized":f"桩号：{pile or '未识别'}{normalization_note}；置信度：{vision.get('confidence','')}；定位：{vision.get('_vision_strategy','original')}；{vision.get('note','')}","_source_text":msg.get("text") or "","_source_group":"Operation","_pile_source_raw":raw_pile,"_pile_normalized":pile!=raw_pile})
            else:
                raw_pile=task["raw_pile"]; pile=normalize_pile_no(raw_pile)
                try: raw_num=float(vision.get("red_number")) if vision.get("red_number") is not None else None
                except (TypeError,ValueError): raw_num=None
                content=raw_num/2 if raw_num is not None else ""
                normalization_note=f"；桩号标准化自 {raw_pile}" if pile!=raw_pile else ""
                records.append({"Date":run["report_date"],"Pile NO.":pile,"Equipment NO.":"","Drilling Start Time":"","Drilling Complete Time":"","Point Complete Time":"","Operator":"","Cement Content":content,"_message_time":msg.get("message_time"),"_image_path":path,"_image_hash":fingerprint,"_meter_reading":raw_num,"_vision_confidence":vision.get("confidence"),"_vision_strategy":vision.get("_vision_strategy","original"),"_recognized":f"仪表读数：{raw_num if raw_num is not None else '未识别'}；写入水泥用量：{content if content != '' else '待人工核对'}{normalization_note}；置信度：{vision.get('confidence','')}；定位：{vision.get('_vision_strategy','original')}","_source_text":task["pile_source"],"_source_group":"Silo","_pile_source_raw":raw_pile,"_pile_normalized":pile!=raw_pile})
        if errors and not records and tasks:
            raise DeepSeekError(f"所有 {total} 张图片识别均失败。首个错误：{errors[0]}")
        run[stage]=records; run["raw"][stage]=raw
        persist=DATA_DIR/run["run_id"]/stage; persist.mkdir(parents=True,exist_ok=True)
        for row in records:
            old=row.get("_image_path")
            if old and os.path.exists(old):
                dest=persist/(uuid.uuid4().hex+Path(old).suffix.lower()); shutil.copy2(old,dest); row["_image_path"]=str(dest)
        shutil.rmtree(temp,ignore_errors=True)
        warnings=build_warnings(run)
        if errors: warnings.append({"type":"missing","title":"部分图片识别失败","detail":f"{len(errors)} 张图片未能识别；首个错误：{errors[0]}"})
        run["artifacts"][stage]=True
        message=f"已处理 {total} 张图片，识别到 {len(records)} 条目标记录。"
        if errors: message+=f"另有 {len(errors)} 张识别失败，请查看预警。"
        return {"count":len(records),"rows":safe_rows(run),"stage_records":records,"warnings":warnings,"analysis":analysis_for(run),"message":message,"failed_count":len(errors)}
    except Exception:
        shutil.rmtree(temp,ignore_errors=True)
        raise


def _queue_media_job(run:dict[str,Any],stage:str,payload:bytes) -> str:
    run_id=run["run_id"]
    with JOB_LOCK:
        active=next((j for j in JOBS.values() if j.get("run_id")==run_id and j.get("stage")==stage and j.get("status") in ("queued","running")),None)
        if active: raise ValueError("这个阶段已有图片识别任务正在运行，请等待它完成后再重试。")
        job_id=uuid.uuid4().hex
        JOBS[job_id]={"job_id":job_id,"run_id":run_id,"stage":stage,"status":"queued","completed":0,"total":0,"message":"已收到文件，准备开始识别。","result":None,"error":None}

    def worker():
        def progress(done,total,last_name):
            message=f"识别进度 {done}/{total} 张"
            if last_name and done: message+=f"；刚完成 {last_name}"
            elif total: message+="；其余图片正在识别"
            with JOB_LOCK:
                job=JOBS.get(job_id)
                if job:
                    job.update(status="running",completed=done,total=total,message=message)
        with JOB_LOCK:
            JOBS[job_id].update(status="running",message="正在解压并检查聊天记录和图片。")
        try:
            result=_extract_media_zip(run,io.BytesIO(payload),stage,on_progress=progress)
            with JOB_LOCK:
                JOBS[job_id].update(status="completed",completed=JOBS[job_id]["total"],message=result["message"],result=result)
        except Exception as exc:
            with JOB_LOCK:
                JOBS[job_id].update(status="failed",message="识别失败。",error=str(exc))

    thread=threading.Thread(target=worker,name=f"media-{stage}-{job_id[:8]}",daemon=True)
    thread.start()
    return job_id


@app.get("/api/jobs/<job_id>")
def media_job_status(job_id):
    with JOB_LOCK:
        job=JOBS.get(job_id)
        if not job: return jsonify({"error":"识别任务不存在或服务已重启，请重新上传。"}),404
        return jsonify({k:v for k,v in job.items() if k not in ("job_id","run_id")})


@app.post("/api/runs/<run_id>/operation")
def upload_operation(run_id):
    try:
        run=get_run(run_id); file=request.files.get("file")
        if not file or not file.filename.lower().endswith(".zip"): return jsonify({"error":"请上传包含 WhatsApp TXT 和图片的 ZIP 文件。"}),400
        job_id=_queue_media_job(run,"operation",file.read())
        return jsonify({"job_id":job_id,"message":"文件已上传，后台开始识别。"}),202
    except (ValueError,DeepSeekError) as e: return jsonify({"error":str(e)}),400
    except Exception as e: return jsonify({"error":f"Operation 处理失败：{e}"}),500


@app.post("/api/runs/<run_id>/silo")
def upload_silo(run_id):
    try:
        run=get_run(run_id); file=request.files.get("file")
        if not file or not file.filename.lower().endswith(".zip"): return jsonify({"error":"请上传包含 WhatsApp TXT 和图片的 ZIP 文件。"}),400
        job_id=_queue_media_job(run,"silo",file.read())
        return jsonify({"job_id":job_id,"message":"文件已上传，后台开始识别。"}),202
    except (ValueError,DeepSeekError) as e: return jsonify({"error":str(e)}),400
    except Exception as e: return jsonify({"error":f"Silo 处理失败：{e}"}),500


@app.get("/api/runs/<run_id>/preview")
def preview(run_id):
    try:
        run=get_run(run_id); return jsonify({"rows":safe_rows(run),"warnings":build_warnings(run),"analysis":analysis_for(run),"report_date":run["report_date"]})
    except ValueError as e: return jsonify({"error":str(e)}),400


@app.post("/api/runs/<run_id>/preview")
def update_preview(run_id):
    try:
        run=get_run(run_id); rows=request.get_json(force=True) or []
        originals=safe_rows(run); run["manual"]=[]
        for index,row in enumerate(rows):
            old=originals[index] if index<len(originals) else None
            old_pile=old.get("Pile NO.") if old else None
            current_pile=normalize_pile_no(row.get("Pile NO."))
            record={field:row.get(field,"") for field in FIELDS}
            record["Pile NO."]=current_pile
            record["_start_dt"]=parse_clock(str(record.get("Drilling Start Time") or ""),str(record.get("Date") or run["report_date"]))[1]
            record["_drill_dt"]=parse_clock(str(record.get("Drilling Complete Time") or ""),str(record.get("Date") or run["report_date"]))[1]
            record["_complete_dt"]=parse_clock(str(record.get("Point Complete Time") or ""),str(record.get("Date") or run["report_date"]))[1]
            sources=[*run["construction"].values(),run["operation"],run["silo"]]
            matched=False
            for source in sources:
                for item in source:
                    if old_pile and normalize_pile_no(item.get("Pile NO."))==normalize_pile_no(old_pile):
                        item.update(record); matched=True
            if not matched and current_pile: run["manual"].append(record)
        return jsonify({"ok":True,"warnings":build_warnings(run),"analysis":analysis_for(run)})
    except ValueError as e: return jsonify({"error":str(e)}),400


@app.get("/api/runs/<run_id>/download/construction.xlsx")
def download_construction_xlsx(run_id):
    try:
        run=get_run(run_id)
        if not run["artifacts"].get("construction"): return jsonify({"error":"请先在施工阶段核对并确认所选 TXT 文件。"}),409
        rows=[r for group in run["construction"].values() for r in group]
        rows.sort(key=lambda r:(r.get("_complete_dt") or "9999",r.get("Pile NO.") or ""))
        return send_file(_write_xlsx(rows),as_attachment=True,download_name=f"施工审核_{run['report_date']}.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except ValueError as e: return jsonify({"error":str(e)}),400


@app.get("/api/runs/<run_id>/download/construction.zip")
def download_construction_zip(run_id):
    try:
        run=get_run(run_id)
        if not run["artifacts"].get("construction"): return jsonify({"error":"请先在施工阶段核对并确认所选 TXT 文件。"}),409
        return send_file(_write_txt_zip(run),as_attachment=True,download_name=f"施工结构化审核材料_{run['report_date']}.zip",mimetype="application/zip")
    except ValueError as e: return jsonify({"error":str(e)}),400


@app.get("/api/runs/<run_id>/download/<stage>.docx")
def download_stage_docx(run_id,stage):
    try:
        run=get_run(run_id)
        if stage not in ("operation","silo"): return jsonify({"error":"审核材料阶段无效。"}),400
        if not run["artifacts"].get(stage): return jsonify({"error":"请先核对并确认本阶段 ZIP 文件，等待处理完成后再下载。"}),409
        records=run[stage]
        heading="Operation 图片审核材料" if stage=="operation" else "Silo Operation 图片审核材料"
        return send_file(_write_review_docx(records,heading,run["report_date"],include_operation_details=(stage=="operation")),as_attachment=True,download_name=f"{stage}_审核材料_{run['report_date']}.docx",mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except ValueError as e: return jsonify({"error":str(e)}),400


@app.get("/api/roster")
def roster(): return jsonify({"roster":load_roster()})


@app.post("/api/roster")
def upload_roster():
    file=request.files.get("file")
    if not file or not file.filename.lower().endswith((".xlsx",".xlsm")): return jsonify({"error":"请上传包含 Account Name 和 Full Name 两列的 Excel 文件（.xlsx）。"}),400
    try:
        wb=load_workbook(file,read_only=True,data_only=True); ws=wb.active
        headers=[str(c.value or "").strip().casefold() for c in ws[1]]
        account_idx=next((i for i,h in enumerate(headers) if "account" in h),None)
        full_idx=next((i for i,h in enumerate(headers) if "full" in h),None)
        if account_idx is None or full_idx is None: return jsonify({"error":"Excel 表头必须包含 Account Name 和 Full Name。"}),400
        result=[]
        for values in ws.iter_rows(min_row=2,values_only=True):
            account=str(values[account_idx] or "").strip(); name=str(values[full_idx] or "").strip()
            if account and name: result.append({"account":account,"full_name":name})
        if not result: return jsonify({"error":"Excel 中没有有效的账户映射数据。"}),400
        ROSTER_PATH.write_text(json.dumps({"roster":result},ensure_ascii=False,indent=2),encoding="utf-8")
        return jsonify({"roster":result,"count":len(result)})
    except Exception as e: return jsonify({"error":f"读取 Excel 失败：{e}"}),400


@app.get("/api/runs/<run_id>/analysis")
def analysis(run_id):
    try: run=get_run(run_id); return jsonify({"analysis":analysis_for(run)})
    except ValueError as e: return jsonify({"error":str(e)}),400


@app.post("/api/runs/<run_id>/write-feishu")
def write_feishu(run_id):
    try: run=get_run(run_id)
    except ValueError as e: return jsonify({"error":str(e)}),400
    rows=safe_rows(run)
    if not rows: return jsonify({"ok":False,"message":"预览表中没有可写入的桩号记录。"}),400
    try:
        requests,domain,token,app_token,table_id=_feishu_context()
        names=sorted({str(r.get("Operator")) for r in rows if r.get("Operator")})
        operator_field=_operator_field(requests,domain,token,app_token,table_id)
        if operator_field and (operator_field.get("type")==3 or operator_field.get("ui_type")=="SingleSelect"):
            have={str(x.get("name","")).casefold() for x in (operator_field.get("property") or {}).get("options",[])}
            missing=[name for name in names if name.casefold() not in have]
            if missing:
                return jsonify({"ok":False,"requires_option_creation":True,"options":missing,"message":"这些操作员全名尚未出现在飞书 Operator 单选选项中。"}),409
        def date_ms(text): return int(datetime.combine(date.fromisoformat(text),datetime.min.time()).timestamp()*1000) if text else None
        def time_ms(text, full):
            value=full or (f"{run['report_date']} {text}" if text else "")
            if not value: return None
            dt=datetime.fromisoformat(value.replace(" ","T") if "T" not in value else value)
            return int(dt.timestamp()*1000)
        api_rows=[]
        for row in rows:
            fields={"Date":date_ms(row.get("Date")),"Pile NO.":row.get("Pile NO."),"Equipment NO.":row.get("Equipment NO."),
                    "Drilling Start Time":time_ms(row.get("Drilling Start Time"),row.get("_start_dt")),
                    "Drilling Complete Time":time_ms(row.get("Drilling Complete Time"),row.get("_drill_dt")),
                    "Point Complete Time":time_ms(row.get("Point Complete Time"),row.get("_complete_dt")),
                    "Operator":row.get("Operator") or None,"Cement Content":float(row["Cement Content"]) if row.get("Cement Content") not in (None,"") else None}
            api_rows.append({"fields":{k:v for k,v in fields.items() if v is not None and (k=="Pile NO." or v!="")}})
        headers={"Authorization":f"Bearer {token}"}
        list_url=f"{domain}/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records"
        existing_by_pile:dict[str,str]={}; duplicate_piles:set[str]=set(); page_token=""
        while True:
            params={"page_size":500}
            if page_token: params["page_token"]=page_token
            res=requests.get(list_url,headers=headers,params=params,timeout=30)
            res.raise_for_status(); payload=res.json()
            if payload.get("code")!=0: raise RuntimeError(payload.get("msg","读取飞书现有记录失败"))
            data=payload.get("data") or {}
            for item in data.get("items") or data.get("records") or []:
                value=(item.get("fields") or {}).get("Pile NO.")
                if isinstance(value,list):
                    value="".join(str(part.get("text") or part.get("name") or "") if isinstance(part,dict) else str(part) for part in value)
                elif isinstance(value,dict): value=value.get("text") or value.get("name") or ""
                pile=normalize_pile_no(value)
                if not pile: continue
                if pile in existing_by_pile: duplicate_piles.add(pile)
                else: existing_by_pile[pile]=str(item.get("record_id") or "")
            if not data.get("has_more"): break
            page_token=str(data.get("page_token") or "")
            if not page_token: break
        target_piles={str(row.get("fields",{}).get("Pile NO.") or "").strip() for row in api_rows}
        conflicts=sorted(pile for pile in duplicate_piles if pile in target_piles)
        if conflicts:
            return jsonify({"ok":False,"message":f"飞书中这些桩号存在重复行，无法确定要覆盖哪一行：{'、'.join(conflicts)}。请先人工合并重复记录。"}),409
        creates=[]; updates=[]
        for item in api_rows:
            pile=str(item["fields"].get("Pile NO.") or "").strip()
            record_id=existing_by_pile.get(pile)
            if record_id: updates.append({"record_id":record_id,"fields":item["fields"]})
            else: creates.append(item)
        create_url=f"{list_url}/batch_create"; update_url=f"{list_url}/batch_update"
        created=updated=0
        for offset in range(0,len(creates),500):
            chunk=creates[offset:offset+500]
            if not chunk: continue
            res=requests.post(create_url,headers=headers,json={"records":chunk},timeout=30)
            res.raise_for_status(); payload=res.json()
            if payload.get("code")!=0: raise RuntimeError(payload.get("msg","飞书新增记录失败"))
            created+=len(payload.get("data",{}).get("records",[])) or len(chunk)
        for offset in range(0,len(updates),500):
            chunk=updates[offset:offset+500]
            if not chunk: continue
            res=requests.post(update_url,headers=headers,json={"records":chunk},timeout=30)
            res.raise_for_status(); payload=res.json()
            if payload.get("code")!=0: raise RuntimeError(payload.get("msg","飞书更新记录失败"))
            updated+=len(payload.get("data",{}).get("records",[])) or len(chunk)
        return jsonify({"ok":True,"created":created,"updated":updated,"written":created+updated,"message":f"飞书同步完成：新增 {created} 条，更新 {updated} 条；已有桩号不会重复创建。"})
    except Exception as e: return jsonify({"ok":False,"message":f"飞书写入失败：{e}"}),502


@app.post("/api/feishu/operator-options")
def create_operator_options():
    body=request.get_json(force=True) or {}; options=body.get("options") or []
    options=[str(name).strip() for name in options if str(name).strip()]
    if not options: return jsonify({"error":"没有需要新增的操作员名称。"}),400
    try:
        added=_add_operator_options(options)
        return jsonify({"ok":True,"added":added,"message":f"已在飞书 Operator 单选字段中新增 {len(added)} 个选项。"})
    except Exception as e: return jsonify({"ok":False,"message":f"添加飞书单选选项失败：{e}"}),502


@app.errorhandler(413)
def too_large(_): return jsonify({"error":f"文件超过 {MAX_UPLOAD//1024//1024} MB 上限。"}),413


if __name__=="__main__":
    app.run(host="127.0.0.1",port=int(os.getenv("PORT","8766")),debug=False)
