from __future__ import annotations

import shutil
import zipfile
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "三类群聊关联测试数据_2026-09-19_账户映射修正版"
OP_SOURCE = ROOT / "source_review" / "operation" / "群聊2图片"
SILO_SOURCE = ROOT / "source_review" / "silo" / "群聊3图片"
ROSTER = Path(r"C:\Users\RANYI XIE\Desktop\names.xlsx")


def whatsapp(timestamp: str, sender: str, text: str) -> str:
    return f"[{timestamp}] {sender}: {text}\n"


def construction_record(pile: str, when: str, start: str, drilled: str, complete: str) -> str:
    return whatsapp(
        when,
        "Site Report",
        "Project : T5 Megaspine\n"
        "Company : KB (Geoharbour)\n"
        f"Date :{when[:10].replace('-', '/')}\n"
        "Location : DSM open cut South area 5A ZONE: E-E\n"
        "KB GH:01\n"
        f"Point No: {pile}\n"
        "RFI no: MSP_CR-RFIw-TEST\n\n"
        "Ground level: +5.85mSHD\n"
        "Main Grouting Top L:-7.00mSHD\n"
        "Main Grouting Toe L:-25.00mSHD\n\n"
        f"Drilling start time:{start}hrs\n"
        f"Drilling complete time:{drilled}hrs\n"
        "Drilling depth:30.85 m\n"
        "Grouting density:1.51kg/L\n"
        "Main Grouting start time:08:12hrs\n"
        "Main Grouting end time:08:40hrs\n"
        f"Point complete time:{complete}hrs\n"
        "Total volume:16,420L",
    )


ROWS = [
    # pile, equipment, drilling start, drilling complete, point complete, report time
    ("D-E-0158", "DCM-1", "12:10", "13:00", "13:33", "2026-09-19 13:34:00"),
    ("G12-R1-16", "DCM-2", "11:37", "12:55", "13:05", "2026-09-19 13:05:00"),
    ("D-E-0155", "DCM-3", "22:55", "23:40", "23:55", "2026-09-19 23:56:00"),
    ("F1-E-0109", "DCM-4", "12:00", "12:50", "13:05", "2026-09-19 13:06:00"),
    ("31-WB-13", "DCM-5", "11:30", "12:20", "12:40", "2026-09-19 12:40:00"),
    ("E-W-0026", "DCM-6", "11:15", "12:15", "12:40", "2026-09-19 12:48:00"),
    ("D-E-0499", "DCM-7", "23:05", "23:50", "00:01", "2026-09-20 00:04:00"),
    ("G12-R1-13", "DCM-8", "05:50", "06:40", "06:55", "2026-09-20 06:58:00"),
    ("D-E-0157", "DCM-1", "09:00", "09:45", "10:00", "2026-09-19 10:00:00"),
    ("G12-R1-15", "DCM-2", "10:00", "10:45", "11:03", "2026-09-19 11:03:00"),
    ("B1-W-A-155AB", "DCM-3", "07:00", "07:08", "07:10", "2026-09-19 07:10:00"),
    ("E-W-031", "DCM-4", "05:00", "05:45", "06:00", "2026-09-20 06:00:00"),
    ("B1-W-A-155C", "DCM-5", "05:00", "05:50", "06:00", "2026-09-20 06:00:00"),
]


# Operation sender labels are synthetic test accounts and must exactly match
# the Account Name values from the supplied names.xlsx roster.
OP_MESSAGES = [
    ("2026-09-19 13:34:00", "Raj", "DSM control screen photo. <attached: media/微信图片_2026-09-26_044603_699.png>"),
    ("2026-09-19 13:06:00", "Dev", "<attached: media/微信图片_2026-09-26_044610_609.png>"),
    ("2026-09-19 13:05:00", "Arjun", "<attached: media/微信图片_2026-09-26_044614_242.png>"),
    ("2026-09-19 12:40:00", "Vikram", "<attached: media/微信图片_2026-09-26_044618_108.png>"),
    ("2026-09-19 12:48:00", "Aditya", "<attached: media/微信图片_2026-09-26_044637_310.png>"),
    ("2026-09-19 23:56:00", "Rohit", "<attached: media/微信图片_2026-09-26_044733_230.png>"),
    ("2026-09-20 00:04:00", "Karan", "<attached: media/微信图片_2026-09-26_044737_460.png>"),
    ("2026-09-20 06:58:00", "Sidd", "<attached: media/微信图片_2026-09-26_044742_400.png>"),
    ("2026-09-19 13:27:00", "Raj", "DSM screen photo; generated message metadata. <attached: media/Codex 图像 2026年9月26日 04_48_34.png>"),
    ("2026-09-19 13:28:00", "Arjun", "Duplicate-pile test photo. <attached: media/微信图片_20260926044534.png>"),
]

SILO_MESSAGES = [
    ("2026-09-19 11:06:00", "Nazrul (Silo 5 operator)", "Silo 5\nPile ID G12-R1-15\nCement mixing finished\n<attached: media/微信图片_2026-09-26_045933_921.png>"),
    ("2026-09-19 10:07:00", "Rajen Raja (Silo 4 operator)", "Silo 3\nPile Id D-E-0157\nCement mixing complete\n<attached: media/微信图片_2026-09-26_045958_242.png>"),
    ("2026-09-19 07:12:00", "Anwar (Silo 2 operator)", "Silo 2 pile ID B1-W-A-155AB\nCement mixing finish.\n<attached: media/微信图片_2026-09-26_050131_310.png>"),
    ("2026-09-20 06:16:00", "Shahadot Islam (Silo 7 operator)", "Silo-8\nPile ID E-W-031\nCement mixing finished\n<attached: media/微信图片_2026-09-26_050135_492.png>"),
    ("2026-09-20 06:14:00", "Anwar (Silo 2 operator)", "Silo 2 pile ID B1-W-A-155C\nCement mixing finish.\n<attached: media/微信图片_2026-09-26_050140_852.png>"),
]


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8-sig", newline="\n")


def make_zip(source_dir: Path, archive_path: Path) -> None:
    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(source_dir.rglob("*")):
            if path.is_file():
                zf.write(path, path.relative_to(source_dir).as_posix())


def build() -> None:
    if OUT.exists():
        raise FileExistsError(f"为避免覆盖已有资料，输出目录已存在：{OUT}")
    construction = OUT / "01_施工审核群聊"
    operation = OUT / "02_Operation群聊"
    silo = OUT / "03_Silo_Operation群聊"
    (construction / "txt").mkdir(parents=True)
    (operation / "media").mkdir(parents=True)
    (silo / "media").mkdir(parents=True)

    per_device: dict[str, list[str]] = {f"DCM-{i}": [] for i in range(1, 9)}
    for pile, equipment, start, drilled, complete, when in ROWS:
        per_device[equipment].append(construction_record(pile, when, start, drilled, complete))
    for device, messages in per_device.items():
        raw = "此聊天以端到端加密。\n" + "".join(messages)
        write_text(construction / "txt" / f"{device}.txt", raw)

    op_chat = "此聊天以端到端加密。\n"
    for timestamp, sender, body in sorted(OP_MESSAGES):
        op_chat += whatsapp(timestamp, sender, body)
    write_text(operation / "WhatsApp Chat - Operation Group.txt", op_chat)
    for image in sorted(OP_SOURCE.glob("*")):
        if image.is_file():
            shutil.copy2(image, operation / "media" / image.name)

    silo_chat = "此聊天以端到端加密。\n"
    for timestamp, sender, body in sorted(SILO_MESSAGES):
        silo_chat += whatsapp(timestamp, sender, body)
    write_text(silo / "WhatsApp Chat - Silo Operation Group.txt", silo_chat)
    for image in sorted(SILO_SOURCE.glob("*")):
        if image.is_file():
            shutil.copy2(image, silo / "media" / image.name)

    guide = """三类群聊关联测试数据 · 2026-09-19

日报日期选择：2026-09-19（时间窗：2026-09-19 07:00 至 2026-09-20 07:00）

上传方式：
1. 施工审核群聊：解压 01_施工审核群聊.zip，选择当天有记录的 DCM-1 至 DCM-8 TXT；没有记录的设备不必上传。
2. Operation 群聊：直接上传 02_Operation群聊.zip。
3. Silo Operation 群聊：直接上传 03_Silo_Operation群聊.zip。
4. 名称映射使用随包提供的 names.xlsx。Operation TXT 的发送账号只使用 Account Name 列中的 Raj、Dev、Arjun、Vikram、Aditya、Rohit、Karan、Sidd，因此可按 Full Name 列匹配预览表 Operator。

数据关联：
- 13 个施工桩号覆盖 Operation 的 8 个桩号及 Silo 的 5 个桩号，供系统按 Pile NO. 合并。
- Silo 仪表红色读数预期分别为 10、33、6、21、3；写入 Cement Content 前除以 2。
- Operation 的 G12-R1-16 被安排了多张图片，用于检查同桩多图/重复上传警示；两张无可见 WhatsApp 发件人信息的单图使用 Demo-* 合成发件人。
- Operation TXT 中的账号名是为完整测试映射流程而使用的合成账号，严格来自 names.xlsx 左列；Agent 应以 TXT 的发送账号映射右列全名。原始图片内部可能显示不同真实群聊名称，图片本身不修改，识别只读取目标机器桩号。

时间说明：
- 三类 TXT 的 WhatsApp 消息头是为同一日报窗口合成的标准导出格式，确保可测试窗口过滤和跨群合并。
- 原始图片未修改；图片里嵌入的拍摄时间可能与合成的 WhatsApp 消息时间不同。系统应以 TXT 消息头时间作为发送时间。
- 这是一套演示输入，所有施工记录中的字段、时间、RFI 号及合成消息头仅供测试，不代表真实日报。
"""
    write_text(OUT / "使用说明.txt", guide)

    make_zip(construction, OUT / "01_施工审核群聊.zip")
    make_zip(operation, OUT / "02_Operation群聊.zip")
    make_zip(silo, OUT / "03_Silo_Operation群聊.zip")

    combined = OUT / "完整测试包"
    combined.mkdir()
    for filename in ("01_施工审核群聊.zip", "02_Operation群聊.zip", "03_Silo_Operation群聊.zip", "使用说明.txt"):
        shutil.copy2(OUT / filename, combined / filename)
    if ROSTER.exists():
        shutil.copy2(ROSTER, combined / "names.xlsx")
    else:
        write_text(combined / "names.xlsx缺失说明.txt", "原 names.xlsx 未在预期桌面位置找到；请将你提供的 names.xlsx 放在本目录。")
    make_zip(combined, OUT / "完整测试包_三类群聊.zip")

    print(f"输出目录：{OUT}")
    for p in sorted(OUT.glob("*.zip")):
        with zipfile.ZipFile(p) as zf:
            print(f"{p.name}: {len([n for n in zf.namelist() if not n.endswith('/')])} 个文件")
    print(f"施工记录：{len(ROWS)}；Operation 原图：{len(list((operation / 'media').iterdir()))}；Silo 原图：{len(list((silo / 'media').iterdir()))}")


if __name__ == "__main__":
    build()
