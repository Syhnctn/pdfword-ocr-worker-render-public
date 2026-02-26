# OCR Worker

FastAPI worker that processes queued jobs from Supabase and uploads markdown/docx output to storage.

Capabilities:
- Extracts embedded text from text-based PDFs locally (`pypdf`) and writes readable DOCX output.
- If configured, can call external OCR (`LIGHTON_OCR_ENDPOINT`) for scanned/image PDFs.

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
- `LIGHTON_OCR_ENDPOINT`
- `LIGHTON_OCR_TOKEN`
