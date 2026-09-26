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
def _dump(sections) -> list[list[str]]:
    return [[int(index), str(text)] for index, text in (sections or [])]


def _write_result(output_path: str, payload: dict) -> None:
    with open(output_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False)


def main() -> int:

    pdf_path, output_path = sys.argv[1], sys.argv[2]

    # Imported here (not at module import time) so the parent can spawn this
    # without paying for the OCR stack twice.
    import main as worker

    try:
        pdf_bytes = open(pdf_path, "rb").read()  # noqa: SIM115 - short-lived
    except Exception as exc:  # noqa: BLE001 - reported back to the parent
        _write_result(output_path, {"ok": False, "error": f"read_failed:{exc}"})
        return 1

    # OCRmyPDF pulls in Ghostscript, which is the heaviest consumer on the
    # 512 MB free tier. Inside a page-sized chunk the plain Tesseract path
    # produces the same text with a fraction of the memory, so it is tried
    # first and OCRmyPDF remains only as a fallback.
    try:
        sections = worker.extract_pdf_text_sections(pdf_bytes)
        if sections:
            _write_result(
                output_path,
                {"ok": True, "sections": _dump(sections), "note": ""},
            )
            return 0
    except Exception:
        pass

    try:
        sections = worker.extract_pdf_text_sections_with_tesseract(pdf_bytes)
        note = (
            "Extracted with open-source OCR (Tesseract fallback)."
            if sections
            else ""
        )
        _write_result(
            output_path, {"ok": True, "sections": _dump(sections), "note": note}
        )
        return 0
    except Exception as exc:  # noqa: BLE001 - reported back to the parent
        _write_result(
            output_path,
            {"ok": False, "error": str(exc), "sections": [], "note": ""},
        )
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
