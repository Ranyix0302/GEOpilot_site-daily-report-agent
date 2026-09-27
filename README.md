# Site Daily Report Agent

WhatsApp 施工日报 Agent。网页使用已确认的四阶段流程；比赛 LLM Gateway（或本机 DeepSeek）及飞书凭证均由后端从 `.env` 文件读取。

## 先准备环境

需要 Windows 10/11、Python 3.11 或更新版本。先在 DeepSeek 开放平台创建 API Key：<https://platform.deepseek.com/>。第一次启动脚本会创建并打开 `.env` 文件：

```powershell
.\start.ps1
```

在 `.env` 中填写 DeepSeek Key 和飞书配置，保存后再次运行 `start.ps1`。部署到比赛 AWS 环境时，改用主办方提供的三个 `LLM_GATEWAY_*` 配置：

```dotenv
DEEPSEEK_API_KEY=你的 DeepSeek API Key
DEEPSEEK_MODEL=deepseek-flash
DEEPSEEK_BASE_URL=https://api.deepseek.com
# AWS 比赛部署使用以下三项，并删除或忽略上面的 DeepSeek 配置
LLM_GATEWAY_URL=https://api.softwaresystems.app
LLM_GATEWAY_API_KEY=主办方发给团队的 API Key
LLM_MODEL=global.anthropic.claude-sonnet-4-5-20250929-v1:0
VISION_CONFIDENCE_THRESHOLD=0.75
LARK_APP_ID=飞书自建应用的 App ID
LARK_APP_SECRET=飞书自建应用的 App Secret
LARK_DOMAIN=https://open.feishu.cn
LARK_BASE_APP_TOKEN=目标多维表格 URL 中的 app_token
LARK_TABLE_ID=目标数据表 URL 中的 table_id
```

不要在 Key 两边加引号，也不要把 `.env` 发给别人或提交到 Git。Agent 只从后端读取 API Key；密钥不会进入前端代码。只要设置了 `LLM_GATEWAY_URL`，系统就优先使用比赛 Gateway；否则沿用本机 DeepSeek 配置。比赛 Gateway 按主办方 Starter Kit 使用 Ollama 兼容的 `/api/chat` 协议和 `X-API-Key` 请求头，图片以消息的 `images` 字段发送。

图片识别先在本机自动准备候选区域：Operation 照片优先识别控制屏右上桩号区域，Silo 照片优先定位红色 LED 区域。候选区域无法得到有效结果时自动回退到原图。置信度低于 `VISION_CONFIDENCE_THRESHOLD` 的记录会进入预警，要求工程师在审核 Word 中对照原图。

## 启动

运行 `start.ps1`，然后打开 <http://127.0.0.1:8766/>。停止时在运行窗口按 `Ctrl+C`。修改 `.env` 后需重启服务。

打开 <http://127.0.0.1:8766/>。停止时在运行窗口按 `Ctrl+C`。改 `.env` 后需重启服务。

## 飞书配置

App ID / App Secret 是应用凭证；App Token / Table ID 用来指向具体多维表格和数据表，是额外必需的定位信息。应用还需获得目标多维表格的读取、编辑权限。将实际地址复制进去即可；**不要把这些凭证放到前端代码里**。

飞书数据表需要有这八个列名，且拼写与大小写完全一致：`Date`、`Pile NO.`、`Equipment NO.`、`Drilling Start Time`、`Drilling Complete Time`、`Point Complete Time`、`Operator`、`Cement Content`。日期与三个时间列应使用日期时间字段，`Operator` 使用单选，`Cement Content` 使用数字字段。写入以 `Pile NO.` 为唯一主键：飞书已有桩号时只更新本次预览中有值的字段，没有该桩号时才新增记录，避免重复行。如果飞书中同一桩号已经存在多行，Agent 会停止写入并要求人工合并。操作员全名不在 `Operator` 现有选项中时，页面会先询问工程师；确认后 Agent 才新增单选选项并继续写入。为此飞书应用需要多维表格记录读取、创建、更新和字段管理权限。

## 当前流程

- 第一步：选择日报日期，只为当天实际工作的 DCM 设备选择对应 WhatsApp TXT（至少一台，最多八台）。未选择的设备视为当天未工作，不作为缺失预警。浏览器先显示所选文件名，工程师确认后才一次发送至后端。系统筛选 `[日报日期 07:00, 次日 07:00)`，识别桩号和三个时间，Equipment NO. 取自设备按钮。处理完成后提供按 Point Complete Time 排序的 Excel，以及包含一份总汇总和各工作设备结构化证据 TXT 的审核 ZIP。
- 第二、三步：选择 Operation 或 Silo Operation ZIP 后先核对文件名和大小，点击确认后才上传识别。处理成功前审核 Word 下载保持锁定；Word 使用结构化信息表和原始证据图片，供工程师对照网页预览表人工审核。
- 三类群聊互不锁定、没有固定处理顺序。工程师可以只处理其中一类或两类并同步飞书；后续上传的群聊数据按 `Pile NO.` 合并，并增量更新飞书中的同一行。
- 第二步：上传 Operation ZIP（需含聊天 TXT 和图片）。只将设备控制屏识别为目标图，提取桩号；从发送账号匹配操作员全名。提供包含账号、全名、发送时间、原图和识别说明的 Word 审核材料。
- 第三步：上传 Silo Operation ZIP。从消息文字读取桩号，识别指定仪表的红色数字并除以二，提供 Word 审核材料。
- 最后：合并数据到只有飞书八个字段的可编辑预览表，显示缺失、冲突、时间差预警和操作员工作量/班次统计；确认后写入飞书。
- 操作员映射表可以用右下/侧栏“操作员映射表 · 上传 Excel”替换，表头应含 `Account Name` 与 `Full Name`。

## 资料和安全

- 上传文件只在本机 Flask 服务运行期间处理；上传的图片审核材料会保存在本机 `data/<run_id>/`，便于下载 Word。
- `.env`、名册和处理图片均被 `.gitignore` 排除。不要把 `.env` 分享给工程师或提交到 Git。工程师应访问同一个已配置的 Agent 服务器网页；若把 `.env` 发给每位工程师，他们都能读取和使用共享 Key。
- 当前处理会话数据保存在服务内存中；服务重启后需要重新上传。映射表和已生成图片保存在本机 `data/`。
- 文件上传上限默认 100 MB，可通过 `.env` 中 `MAX_UPLOAD_MB` 修改。
- 微信/WhatsApp 导出格式、图片内容和飞书字段类型仍需用你的脱敏样本实际验证；图片识别不确定时应由工程师按 Word 原图材料核对。

## 目录

```text
agent_v2/
  agent/app.py          Flask API、解析、审核材料、飞书写入
  agent/deepseek_vision.py  DeepSeek 图像连接、自动裁剪和原图兜底
  templates/index.html  网页
  static/               CSS 和前端交互
  data/                 本机名册和临时审核图片
  .env.example          密钥配置模板
```
