# PDF to Word Pro (Flutter UI Prototype)

Dark themed mobile UI prototype based on provided reference screens.

## Implemented

- Login screen with mock Apple/Google/Email auth actions
- Convert screen with real file picker, signed upload, and job polling flow
- History screen with grouped conversion items
- History screen with row status badges and signed URL download action
- Settings screen with profile, toggles, and language switch
- Routing via `go_router` (`/auth`, `/convert`, `/history`, `/settings`)
- State management via `flutter_riverpod`
- TR/EN localization skeleton (TR default)
- Widget/unit test coverage and golden test scaffold

## Local Flutter Tooling (this workspace)

Flutter SDK is installed in:

`c:\projeler\pdfword\tools\flutter`

Use full path commands if `flutter` is not in your shell PATH:

```bash
c:\projeler\pdfword\tools\flutter\bin\flutter.bat pub get
c:\projeler\pdfword\tools\flutter\bin\flutter.bat test
```

Or use wrapper:

```powershell
.\scripts\flutterw.ps1 pub get
.\scripts\flutterw.ps1 test
```

## Run

```bash
flutter pub get
flutter run
```

## Test

```bash
flutter test
```

## Real Backend Mode (Supabase + OCR Worker)

The app supports real backend mode behind compile-time flags:

- `USE_REAL_BACKEND=true`
- `SUPABASE_URL=...`
- `SUPABASE_ANON_KEY=...`

Example:

```bash
flutter run -d chrome --dart-define=USE_REAL_BACKEND=true --dart-define=SUPABASE_URL=https://YOUR.supabase.co --dart-define=SUPABASE_ANON_KEY=YOUR_ANON_KEY
```

### Supabase assets

- SQL migration: `supabase/migrations/20260216190000_ocr_schema.sql`
- Edge functions:
  - `supabase/functions/create_job/index.ts`
  - `supabase/functions/enqueue_job/index.ts`
  - `supabase/functions/get_job_status/index.ts`
  - `supabase/functions/get_download_url/index.ts`
  - `supabase/functions/process_job/index.ts` (worker webhook target)

### OCR worker

- Code: `backend/ocr_worker/main.py`
- Dockerfile: `backend/ocr_worker/Dockerfile`
- Requirements: `backend/ocr_worker/requirements.txt`
- Supports local embedded-text extraction for text PDFs (`pypdf`), open-source OCR fallback (`Tesseract`) for scanned PDFs, and optional external OCR fallback.

Worker env vars:

- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`
- `OCR_INPUT_BUCKET` (default: `ocr-inputs`)
- `OCR_RESULTS_BUCKET` (default: `ocr-results`)
- `OPEN_SOURCE_OCR_ENABLED` (default: `true`)
- `TESSERACT_LANG` (default: `tur+eng`)
- `TESSERACT_DPI` (default: `220`)
- `TESSERACT_PSM` (default: `6`)
- `TESSERACT_OEM` (default: `1`)
- `TESSERACT_PSM_CANDIDATES` (default: `6,4,11`)
- Optional LightOn endpoint integration:
  - `LIGHTON_OCR_ENDPOINT`
  - `LIGHTON_OCR_TOKEN`

## Production notes

- Public functions are deployed with JWT verification enabled:
  - `create_job`, `enqueue_job`, `get_job_status`, `get_download_url`
- Internal worker endpoint:
  - `process_job` is deployed with `--no-verify-jwt` and protected by `OCR_WORKER_SECRET`
- Required Supabase secrets:
  - `OCR_WORKER_WEBHOOK=https://<project-ref>.supabase.co/functions/v1/process_job` (Supabase internal worker)
  - or `OCR_WORKER_WEBHOOK=https://<your-worker-host>/internal/process` (Python worker, recommended for local text extraction)
  - `OCR_WORKER_SECRET=<random-strong-secret>`
