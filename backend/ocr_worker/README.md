# OCR Worker

FastAPI worker that processes queued jobs from Supabase and uploads markdown/docx output to storage.

Capabilities:
- Extracts embedded text from text-based PDFs locally (`pypdf`) and writes readable DOCX output.
- Tries open-source OCR (`OCRmyPDF + Tesseract`) for scanned/image PDFs, then falls back to direct Tesseract OCR.
- If configured, can call external OCR (`LIGHTON_OCR_ENDPOINT`) as an additional fallback.
- Render Docker image is tuned for Turkish OCR by installing `tessdata_best` `tur.traineddata` while keeping local `eng`/`osd` models.

## Run locally

```bash
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 8080
```

## Required env

- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`

## Optional env

- `OCR_INPUT_BUCKET` (default: `ocr-inputs`)
- `OCR_RESULTS_BUCKET` (default: `ocr-results`)
- `OCR_WORKER_SECRET` (if your webhook requires bearer auth)
- `OPEN_SOURCE_OCR_ENABLED` (default: `true`)
- `OCRMYPDF_ENABLED` (default: `true`)
- `OCRMYPDF_LANG` (default: uses `TESSERACT_LANG`)
- `OCRMYPDF_JOBS` (default: `1`)
- `OCRMYPDF_TIMEOUT_SEC` (default: `240`)
- `OCRMYPDF_TESSERACT_TIMEOUT_SEC` (default: derived from `TESSERACT_CALL_TIMEOUT_SEC`)
- `OCRMYPDF_FORCE_OCR` (default: `true`)
- `OCRMYPDF_ROTATE_PAGES` (default: `false`)
- `OCRMYPDF_DESKEW` (default: `false`)
- `OCRMYPDF_CLEAN_FINAL` (default: `false`)
- `OCRMYPDF_OUTPUT_TYPE` (default: `pdf`)
- `TESSERACT_LANG` (default: `tur`)
- `TESSERACT_DPI` (default: `300`)
- `TESSERACT_PSM` (default: `6`)
- `TESSERACT_OEM` (default: `1`)
- `TESSERACT_PSM_CANDIDATES` (default: `6,4`)
- `TESSERACT_MAX_VARIANTS` (default: `3`)
- `TESSERACT_CALL_TIMEOUT_SEC` (default: `10`)
- `TESSERACT_MAX_ATTEMPTS` (default: `3`)
- `LIGHTON_OCR_ENDPOINT`
- `LIGHTON_OCR_TOKEN`
