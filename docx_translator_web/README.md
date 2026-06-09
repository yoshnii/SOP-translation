# DOCX 保格式翻译 · 内部网站

上传 Word 文档 → 豆包翻译（中→英）→ 完整保留图片/排版/表格/目录 → 在线预览 + 下载。
内置 NGS 专业术语表（209 条）+ 整句修复（45 条）。翻完显示费用，供使用者结算。

---

## 一、用 Docker 启动（推荐，朋友/Windows 都能用）

需要先装 [Docker Desktop](https://www.docker.com/products/docker-desktop/)（Windows/Mac 都有）。

```bash
# 1. 进入本目录
cd docx_translator_web

# 2. 配置你的火山方舟 API key
cp .env.example .env
#   然后用记事本/编辑器打开 .env，把 ARK_API_KEY 改成你的真实 key

# 3. 一键构建并启动
docker compose up -d --build

# 4. 浏览器打开
#    本机：     http://localhost:8000
#    同事访问： http://<你这台机器的局域网IP>:8000
```

停止：`docker compose down`　查看日志：`docker compose logs -f`

> **谁付费**：服务器用的是 `.env` 里你的 key，所有人翻译都走你的额度。
> 每次翻完页面会显示「本次实际费用 ￥X」，让使用者照此金额转你即可。

---

## 二、不用 Docker，本地直接跑（开发/测试）

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt        # Windows: .venv\Scripts\pip install -r requirements.txt
export ARK_API_KEY=ark-你的key                    # Windows: set ARK_API_KEY=ark-你的key
.venv/bin/uvicorn app:app --host 0.0.0.0 --port 8000
```

---

## 三、功能说明

| 区块 | 功能 |
|---|---|
| ① 上传 | 拖拽或点击上传 .docx |
| ② 分析 | 显示段数、图片数、字符数、**预估费用** |
| ③ 进度 | 实时进度条（按段推进） |
| ④ 结算 | **实际费用**、输入/输出字符、图片保留数、残留中文数；下载 + 在线预览 |
| 术语表 | 只读展示 209 条术语，可搜索；带「纠错」标的含错译纠正规则 |

## 四、修改术语表

术语表是 `glossary.csv`，整句修复是 `overrides.csv`，**只能改文件**（网页只读）。
改完重启服务（Docker：`docker compose up -d --build`）即生效。

格式：
- `glossary.csv`：`中文,英文,备注,错译`（第4列可选，分号分隔豆包的错误译法，会被强制纠正）
- `overrides.csv`：`原文,译文`（原文精确命中时直接用指定译文，优先级最高）

## 五、注意

- 文本框内文字、SmartArt、图表内文字暂不翻译。
- 段内逐字格式（某词单独加粗）会被统一为段落格式。
- 这是机器翻译 + 术语表方案，适合**内部参考/初稿**；正式对外交付建议人工审校。
- 产物存在 `data/` 目录（按 job_id 分文件夹），可定期清理。
