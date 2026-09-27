"""Background job runner for the OCR worker.

``/internal/process`` returns as soon as the job is accepted and hands the work
to this module, which is spawned as a detached process. The free Render tier
closes idle connections after roughly a minute, so a long scan (several minutes
for a multi-page document) cannot stay inside one HTTP request. Running out of
band means the client polls ``get_job_status`` for the result instead.

Called as::

    python -m ocr_job_worker <job_id>
"""

from __future__ import annotations

import asyncio
import sys


def main() -> int:
    if len(sys.argv) != 2:
        print("usage: ocr_job_worker <job_id>", file=sys.stderr)
        return 2

    job_id = sys.argv[1].strip()
    if not job_id:
        return 2

    import main as worker

    try:
        asyncio.run(worker.process_job(job_id))
    except Exception as exc:  # noqa: BLE001 - recorded on the job row
        worker.report_background_job_failure(job_id, exc)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
