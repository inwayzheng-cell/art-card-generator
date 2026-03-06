FROM python:3.11-slim

WORKDIR /app

# 安裝 PDF 處理必備的底層工具
RUN apt-get update && apt-get install -y \
    libpoppler-cpp-dev \
    pkg-config \
    python3-dev \
    && rm -rf /var/lib/apt/lists/*

COPY . .

RUN pip install --no-cache-dir -r requirements.txt

# 使用 7860 端口（Hugging Face 預設）
EXPOSE 7860

CMD ["streamlit", "run", "app.py", "--server.port=7860", "--server.address=0.0.0.0"]