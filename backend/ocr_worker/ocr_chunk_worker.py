"""Isolated OCR runner.

The free Render tier has 512 MB of RAM, which a multi-page scanned PDF blows
through when the whole document is rasterised in one process. This module is
executed as a separate process by :func:`main.run_ocr_in_subprocess`, one small
page group at a time, so the worker survives arbitrarily long documents: the
process exits (freeing all memory) between groups.

Called as::

    python -m ocr_chunk_worker <input.pdf> <output.json>
"""

from __future__ import annotations

import json
import sys


def main() -> int:
    if len(sys.argv) != 3:
        print("usage: ocr_chunk_worker <input.pdf> <output.json>", file=sys.stderr)
        return 2

    input_path, output_path = sys.argv[1], sys.argv[2]

    # Imported here (not at module import time) so the parent can spawn this
    # without paying for the OCR stack twice.
    import main as worker

    with open(input_path, "rb") as handle:
        pdf_bytes = handle.read()

    try:
        sections, note = worker.extract_sections_for_pdf(pdf_bytes)
        payload = {
            "ok": True,
            "sections": [[int(index), str(text)] for index, text in sections],
            "note": str(note or ""),
        }
    except Exception as exc:  # noqa: BLE001 - reported back to the parent
        payload = {"ok": False, "error": str(exc), "sections": [], "note": ""}

    with open(output_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
