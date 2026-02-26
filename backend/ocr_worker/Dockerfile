FROM python:3.12-slim

WORKDIR /app
ENV PYTHONUNBUFFERED=1 \
    LANG=C.UTF-8 \
    LC_ALL=C.UTF-8 \
    TESSDATA_PREFIX=/opt/tessdata-best

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        ghostscript \
        ocrmypdf \
        qpdf \
        tesseract-ocr \
        tesseract-ocr-eng \
        tesseract-ocr-tur \
    && mkdir -p /opt/tessdata-best \
    && cp /usr/share/tesseract-ocr/5/tessdata/eng.traineddata /opt/tessdata-best/eng.traineddata \
    && cp /usr/share/tesseract-ocr/5/tessdata/osd.traineddata /opt/tessdata-best/osd.traineddata \
    && python -c "import urllib.request; urllib.request.urlretrieve('https://raw.githubusercontent.com/tesseract-ocr/tessdata_best/main/tur.traineddata', '/opt/tessdata-best/tur.traineddata')" \
    && python -c "import os; p='/opt/tessdata-best/tur.traineddata'; s=os.path.getsize(p); assert s > 1000000, s; print('tur.traineddata bytes', s)" \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py .

EXPOSE 8080

CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-8080}"]
