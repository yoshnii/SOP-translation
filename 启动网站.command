#!/bin/bash
# SOP 翻译网站启动器
cd "/Users/shan/Documents/BGI/交付中心/SIRO-48_NIFTY翻译/docx_translator_web" || exit 1

# 首次:建虚拟环境 + 装依赖
if [ ! -d ".venv" ]; then
  echo "首次运行,正在安装依赖(约1-2分钟,需联网)..."
  python3 -m venv .venv
  .venv/bin/pip install --quiet -r requirements.txt
fi

# 读 key:.env 优先
if [ -f ".env" ]; then
  export $(grep -v '^#' .env | xargs)
fi
if [ -z "$ARK_API_KEY" ]; then
  echo "⚠️  未找到 ARK_API_KEY。请在 docx_translator_web/.env 里写 ARK_API_KEY=ark-你的key"
  read -p "或现在粘贴key回车(仅本次): " ARK_API_KEY
  export ARK_API_KEY
fi

echo ""
echo "=========================================="
echo "  SOP 翻译网站已启动"
echo "  浏览器打开:  http://localhost:8000"
echo "  关闭:        本窗口按 Ctrl+C 或直接关窗口"
echo "=========================================="

( sleep 3 && open "http://localhost:8000" ) &
.venv/bin/uvicorn app:app --host 127.0.0.1 --port 8000
