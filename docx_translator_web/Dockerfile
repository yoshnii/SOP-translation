FROM python:3.12-slim

WORKDIR /app

# 先装依赖(利用层缓存)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 拷贝应用代码与术语表
COPY app.py translator_core.py glossary.csv overrides.csv ./
COPY static ./static

# 运行期产物目录
RUN mkdir -p /app/data

EXPOSE 8000

# ARK_API_KEY 在运行时通过环境变量注入(见 docker-compose.yml / .env)
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
