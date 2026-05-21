FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    ffmpeg \
    libreoffice \
    poppler-utils \
    fonts-liberation \
    fonts-noto \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY embed_audio.py .
COPY content_studio.py .

EXPOSE 10000

CMD ["python", "content_studio.py"]
