# OCR Worker

FastAPI worker that processes queued jobs from Supabase and uploads markdown/docx output to storage.

Capabilities:
- Extracts embedded text from text-based PDFs locally (`pypdf`) and writes readable DOCX output.
- Tries open-source OCR (`Tesseract`) for scanned/image PDFs.
- If configured, can call external OCR (`LIGHTON_OCR_ENDPOINT`) as an additional fallback.

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
- `TESSERACT_LANG` (default: `tur+eng`)
- `TESSERACT_DPI` (default: `180`)
- `TESSERACT_PSM` (default: `6`)
- `TESSERACT_OEM` (default: `1`)
- `TESSERACT_PSM_CANDIDATES` (default: `6,4`)
- `TESSERACT_MAX_VARIANTS` (default: `3`)
- `TESSERACT_CALL_TIMEOUT_SEC` (default: `8`)
- `TESSERACT_MAX_ATTEMPTS` (default: `4`)
- `LIGHTON_OCR_ENDPOINT`
- `LIGHTON_OCR_TOKEN`
