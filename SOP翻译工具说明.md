# SOP 保格式翻译工具

中文 SOP 文档 → 英文,**完整保留图片、logo、排版、表格、目录、页眉页脚**。
专为 NGS 测序 / 分子诊断领域 SOP 设计,术语专业、跨文档统一。

## 两种用法

| | 位置 | 适合 |
|---|---|---|
| **命令行版** | `cli/` | 批量翻译、自己跑 |
| **网站版** | `web/` | 拖文件进网页,带进度/费用/预览/下载 |

---

## 翻译引擎

| 引擎 | 质量 | 成本/篇 | 说明 |
|---|---|---|---|
| **DeepSeek V4-Pro** | ⭐ 最高 | ~¥0.8 | 章节级翻译，上下文消歧，推荐 |
| DeepSeek V4-Flash | 好 | ~¥0.15 | 快、便宜，个别术语需术语表兜 |
| 豆包翻译模型 | 一般 | ~¥0.1 | 最便宜，但专业术语弱 |

引擎都走火山方舟（同一个 `ARK_API_KEY`），模型需在[火山控制台](https://console.volcengine.com/ark)开通。

---

## 命令行版（cli/）

```bash
cd cli
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
export ARK_API_KEY=ark-你的key

# DeepSeek V4-Pro 翻译（推荐，质量最高）
.venv/bin/python deepseek_translator.py 输入.docx -o 输出_EN.docx

# 豆包翻译（便宜）
.venv/bin/python translate_docx.py 输入.docx -o 输出_EN.docx
```

## 网站版（web/）

```bash
cd web
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
export ARK_API_KEY=ark-你的key
.venv/bin/uvicorn app:app --host 127.0.0.1 --port 8000
# 浏览器打开 http://localhost:8000
```

或用 Docker：`cd web && docker compose up -d --build`（先把 key 填进 `.env`）。

---

## 术语表

- `glossary.csv`：核心术语（试剂/耗材/流程名），跨文档统一。`中文,英文,备注[,错译]`
- `overrides.csv`：整句覆盖（人工修复的标准译文），原文精确命中时直接用。`原文,译文`

改完重启服务即生效。cli 和 web 各有一份，改动需同步两边。

---

## 保格式说明

✅ 保留：图片/logo、表格、目录、页眉页脚、自动编号、制表符、换行、段落格式
⚠️ 不保留：图片**内部**的文字（不可编辑）、段内逐字格式（某词单独加粗会被统一）

## 注意

- 模型在火山控制台有「安心体验」额度上限，反复重翻会触发 429 暂停，需调高/转按量付费。
- API key 不要硬编码进代码，用环境变量 `ARK_API_KEY`。
