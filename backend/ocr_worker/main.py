import base64
import io
import os
import re
import signal
import shutil
import subprocess
import tempfile
import time
from datetime import datetime, timezone
from functools import lru_cache
from typing import Any

import httpx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from pydantic import BaseModel
from pypdf import PdfReader
from supabase import Client, create_client

try:
    import pymupdf as fitz  # PyMuPDF (modern package name)
except Exception:  # pragma: no cover - optional dependency
    try:
        import fitz  # legacy PyMuPDF module name
    except Exception:  # pragma: no cover - optional dependency
        fitz = None

try:
    import pytesseract
    from pytesseract import Output
    from PIL import Image
    from PIL import ImageFilter, ImageOps
except Exception:  # pragma: no cover - optional dependency
    pytesseract = None
    Output = None
    Image = None
    ImageFilter = None
    ImageOps = None

app = FastAPI(title="pdfword-ocr-worker", version="0.3.0")

_BULLET_RE = re.compile(r"^([\-*]|\d+[.)])\s+")
_MULTISPACE_RE = re.compile(r"\s+")
_PUNCT_END_RE = re.compile(r"[.!?:;)](?:['\"])?$")
_HEADING_RE = re.compile(r"^[A-Z0-9][A-Z0-9\s/&()_-]{2,}$")
_WORD_RE = re.compile(r"\w+", re.UNICODE)

_PAGEBREAK_MARKER = "<!-- pagebreak -->"

_TURKISH_CHARS = "\u00e7\u011f\u0131\u00f6\u015f\u00fc\u00c7\u011e\u0130\u00d6\u015e\u00dc"


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def env_flag(name: str, default: bool) -> bool:
    raw = os.environ.get(name, "").strip().lower()
    if not raw:
        return default
    return raw not in {"0", "false", "no", "off"}


def open_source_ocr_enabled() -> bool:
    return env_flag("OPEN_SOURCE_OCR_ENABLED", True)


def tesseract_langs() -> str:
    value = os.environ.get("TESSERACT_LANG", "tur").strip()
    return value or "tur"


def tesseract_final_lang_candidates() -> list[str]:
    available = set(pytesseract_languages()) or set(tesseract_cli_languages())
    base = tesseract_langs()
    candidates = [base]
    if "tur" in base and base != "tur":
        candidates.insert(0, "tur")
    if "eng" not in base:
        candidates.append("eng")
    seen: set[str] = set()
    result: list[str] = []
    for item in candidates:
        key = item.strip()
        if not key or key in seen:
            continue
        if available and "+" not in key and key not in available:
            continue
        seen.add(key)
        result.append(key)
    if result:
        return result
    fallbacks = [lang for lang in ("tur", "eng") if not available or lang in available]
    return fallbacks or ["eng"]


def tesseract_dpi() -> int:
    # 300 dpi renders an A4 page as 2481x3507 (~25 MB in RGB). With the derived
    # image variants and Tesseract's own allocations that peaks past the 512 MB
    # free tier on multi-page scans, which killed the worker with 502/503.
    # 200 dpi still reads normal body text and print accurately.
    raw = os.environ.get("TESSERACT_DPI", "200").strip()
    try:
        dpi = int(raw)
    except ValueError:
        dpi = 200
    return max(96, min(dpi, 600))


def tesseract_psm() -> str:
    raw = os.environ.get("TESSERACT_PSM", "6").strip()
    return raw or "6"


def tesseract_oem() -> str:
    raw = os.environ.get("TESSERACT_OEM", "1").strip()
    return raw if raw in {"0", "1", "2", "3"} else "1"


def tesseract_psm_candidates() -> list[str]:
    raw = os.environ.get("TESSERACT_PSM_CANDIDATES", "").strip()
    values = raw.split(",") if raw else [tesseract_psm()]
    seen: set[str] = set()
    result: list[str] = []
    for part in values:
        item = part.strip()
        if not item or not item.isdigit() or item in seen:
            continue
        seen.add(item)
        result.append(item)
    return result or ["6"]


def tesseract_max_variants() -> int:
    raw = os.environ.get("TESSERACT_MAX_VARIANTS", "2").strip()
    try:
        value = int(raw)
    except ValueError:
        value = 2
    return max(1, min(value, 12))


def tesseract_call_timeout_sec() -> float:
    raw = os.environ.get("TESSERACT_CALL_TIMEOUT_SEC", "10").strip()
    try:
        value = float(raw)
    except ValueError:
        value = 10.0
    return max(2.0, min(value, 60.0))


def tesseract_max_attempts() -> int:
    raw = os.environ.get("TESSERACT_MAX_ATTEMPTS", "3").strip()
    try:
        value = int(raw)
    except ValueError:
        value = 3
    return max(1, min(value, 20))


def tesseract_cmd() -> str:
    value = os.environ.get("TESSERACT_CMD", "").strip()
    return value or "tesseract"


def ocrmypdf_enabled() -> bool:
    return env_flag("OCRMYPDF_ENABLED", True)


def ocrmypdf_cmd() -> str:
    value = os.environ.get("OCRMYPDF_CMD", "").strip()
    return value or "ocrmypdf"


def ocrmypdf_langs() -> str:
    value = os.environ.get("OCRMYPDF_LANG", "").strip()
    return value or tesseract_langs()


def ocrmypdf_jobs() -> int:
    raw = os.environ.get("OCRMYPDF_JOBS", "1").strip()
    try:
        value = int(raw)
    except ValueError:
        value = 1
    return max(1, min(value, 8))


def ocrmypdf_timeout_sec() -> float:
    raw = os.environ.get("OCRMYPDF_TIMEOUT_SEC", "240").strip()
    try:
        value = float(raw)
    except ValueError:
        value = 600.0
    return max(30.0, min(value, 3600.0))


def ocrmypdf_tesseract_timeout_sec() -> int:
    raw = os.environ.get("OCRMYPDF_TESSERACT_TIMEOUT_SEC", "").strip()
    if raw:
        try:
            value = int(float(raw))
        except ValueError:
            value = 0
        if value > 0:
            return max(5, min(value, 600))
    return max(10, min(int(tesseract_call_timeout_sec() * 6), 600))


def ocrmypdf_force_ocr() -> bool:
    return env_flag("OCRMYPDF_FORCE_OCR", True)


def ocrmypdf_rotate_pages() -> bool:
    return env_flag("OCRMYPDF_ROTATE_PAGES", False)


def ocrmypdf_deskew() -> bool:
    return env_flag("OCRMYPDF_DESKEW", False)


def ocrmypdf_clean_final() -> bool:
    return env_flag("OCRMYPDF_CLEAN_FINAL", False)


def ocrmypdf_output_type() -> str:
    value = os.environ.get("OCRMYPDF_OUTPUT_TYPE", "pdf").strip().lower()
    return value or "pdf"


def _candidate_tessdata_dirs() -> list[str]:
    candidates: list[str] = []
    env_path = os.environ.get("TESSDATA_PREFIX", "").strip()
    if env_path:
        candidates.append(env_path)
        candidates.append(os.path.join(env_path, "tessdata"))

    candidates.extend(
        [
            "/usr/share/tesseract-ocr/5/tessdata",
            "/usr/share/tesseract-ocr/4.00/tessdata",
            "/usr/share/tesseract-ocr/tessdata",
            "/usr/share/tessdata",
            "/usr/local/share/tessdata",
        ]
    )

    seen: set[str] = set()
    result: list[str] = []
    for item in candidates:
        path = os.path.abspath(item)
        if path in seen:
            continue
        seen.add(path)
        result.append(path)
    return result


def _looks_like_tessdata_dir(path: str) -> bool:
    if not os.path.isdir(path):
        return False
    for name in ("eng.traineddata", "tur.traineddata", "osd.traineddata"):
        if os.path.exists(os.path.join(path, name)):
            return True
    return False


@lru_cache(maxsize=1)
def resolve_tessdata_dir() -> str:
    for path in _candidate_tessdata_dirs():
        if _looks_like_tessdata_dir(path):
            return path
    return ""


def _clear_tessdata_cache() -> None:
    resolve_tessdata_dir.cache_clear()


@lru_cache(maxsize=1)
def ensure_tesseract_runtime_config() -> dict[str, Any]:
    cmd = tesseract_cmd()
    tessdata_dir = resolve_tessdata_dir()

    if pytesseract is not None:
        try:
            pytesseract.pytesseract.tesseract_cmd = cmd
        except Exception:
            pass

    if tessdata_dir:
        os.environ["TESSDATA_PREFIX"] = tessdata_dir

    info = {
        "tesseract_cmd": cmd,
        "tesseract_path": shutil.which(cmd),
        "tessdata_dir": tessdata_dir,
        "tessdata_candidates": _candidate_tessdata_dirs(),
        "eng_traineddata_exists": bool(
            tessdata_dir and os.path.exists(os.path.join(tessdata_dir, "eng.traineddata"))
        ),
        "tur_traineddata_exists": bool(
            tessdata_dir and os.path.exists(os.path.join(tessdata_dir, "tur.traineddata"))
        ),
    }
    print(
        "[ocr-runtime]",
        {
            "tesseract_path": info["tesseract_path"],
            "tessdata_dir": info["tessdata_dir"],
            "tur_traineddata_exists": info["tur_traineddata_exists"],
        },
    )
    return info


def _tesseract_subprocess_env() -> dict[str, str]:
    ensure_tesseract_runtime_config()
    return dict(os.environ)


def _ocrmypdf_subprocess_env() -> dict[str, str]:
    ensure_tesseract_runtime_config()
    return dict(os.environ)


def _parse_lang_lines(raw: str) -> list[str]:
    langs: list[str] = []
    for line in raw.splitlines():
        value = line.strip()
        if not value:
            continue
        if value.lower().startswith("list of available languages"):
            continue
        langs.append(value)
    return langs


def tesseract_cli_languages() -> list[str]:
    cmd = tesseract_cmd()
    try:
        res = subprocess.run(
            [cmd, "--list-langs"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=10,
            env=_tesseract_subprocess_env(),
            check=False,
        )
    except Exception:
        return []
    text_out = (res.stdout or "") + "\n" + (res.stderr or "")
    return _parse_lang_lines(text_out)


def pytesseract_languages() -> list[str]:
    if pytesseract is None:
        return []

    info = ensure_tesseract_runtime_config()
    config = ""
    tessdata_dir = str(info.get("tessdata_dir") or "").strip()
    if tessdata_dir:
        config = f'--tessdata-dir "{tessdata_dir}"'

    try:
        langs = pytesseract.get_languages(config=config)
    except TypeError:
        langs = pytesseract.get_languages()
    except Exception:
        return []

    result = [str(item).strip() for item in (langs or []) if str(item).strip()]
    seen: set[str] = set()
    deduped: list[str] = []
    for item in result:
        if item in seen:
            continue
        seen.add(item)
        deduped.append(item)
    return deduped


def ocr_runtime_debug_info() -> dict[str, Any]:
    info = dict(ensure_tesseract_runtime_config())
    tessdata_dir = str(info.get("tessdata_dir") or "").strip()

    def _traineddata_size(name: str) -> int | None:
        if not tessdata_dir:
            return None
        path = os.path.join(tessdata_dir, name)
        try:
            return os.path.getsize(path)
        except Exception:
            return None

    info.update(
        {
            "env_TESSDATA_PREFIX": os.environ.get("TESSDATA_PREFIX", ""),
            "env_TESSERACT_LANG": os.environ.get("TESSERACT_LANG", ""),
            "env_OCRMYPDF_LANG": os.environ.get("OCRMYPDF_LANG", ""),
            "pytesseract_languages": pytesseract_languages(),
            "tesseract_cli_languages": tesseract_cli_languages(),
            "ocrmypdf_path": shutil.which(ocrmypdf_cmd()),
            "ghostscript_path": shutil.which("gs"),
            "qpdf_path": shutil.which("qpdf"),
            "tur_traineddata_bytes": _traineddata_size("tur.traineddata"),
            "eng_traineddata_bytes": _traineddata_size("eng.traineddata"),
            "osd_traineddata_bytes": _traineddata_size("osd.traineddata"),
        }
    )

    cmd = tesseract_cmd()
    try:
        version = subprocess.run(
            [cmd, "--version"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=10,
            env=_tesseract_subprocess_env(),
            check=False,
        )
        info["tesseract_version"] = (version.stdout or version.stderr or "").splitlines()[:3]
    except Exception as exc:
        info["tesseract_version_error"] = str(exc)

    for binary, key in (
        (ocrmypdf_cmd(), "ocrmypdf_version"),
        ("gs", "ghostscript_version"),
        ("qpdf", "qpdf_version"),
    ):
        try:
            version = subprocess.run(
                [binary, "--version"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=10,
                env=_ocrmypdf_subprocess_env(),
                check=False,
            )
            info[key] = (version.stdout or version.stderr or "").splitlines()[:3]
        except Exception as exc:
            info[f"{key}_error"] = str(exc)
    return info


def make_supabase_client() -> Client:
    url = os.environ.get("SUPABASE_URL", "").strip()
    key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "").strip()
    if not url or not key:
        raise RuntimeError("SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY are required")
    return create_client(url, key)


def assert_worker_secret(request: Request) -> None:
    expected = os.environ.get("OCR_WORKER_SECRET", "").strip()
    if not expected:
        return

    auth = request.headers.get("Authorization", "")
    token = auth.replace("Bearer ", "").strip()
    if not token or token != expected:
        raise HTTPException(status_code=401, detail="invalid_worker_secret")


def sanitize_storage_name(name: str) -> str:
    safe = re.sub(r"[^a-zA-Z0-9._-]", "_", name.strip())
    return safe or "input.pdf"


def coerce_input_files(
    job_id: str, input_meta: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    input_bucket = os.environ.get("OCR_INPUT_BUCKET", "ocr-inputs")
    files: list[dict[str, Any]] = []

    for index, item in enumerate(input_meta):
        if not isinstance(item, dict):
            continue

        name = str(item.get("name", f"input-{index}.pdf"))
        path = str(item.get("path") or item.get("storage_path") or "").strip()
        bucket = str(item.get("bucket") or item.get("storage_bucket") or input_bucket)
        mime_type = str(item.get("mime_type") or "application/pdf")

        if not path:
            legacy_name = sanitize_storage_name(name)
            path = f"{job_id}/input/{index:02d}_{legacy_name}"

        files.append(
            {
                "name": name,
                "bucket": bucket,
                "path": path,
                "mime_type": mime_type,
                "size_mb": item.get("size_mb"),
            }
        )

    return files


def build_placeholder_markdown(input_meta: list[dict[str, Any]]) -> str:
    lines = ["# OCR Result", ""]
    lines.append("This output was generated by backend worker.")
    lines.append("")
    lines.append("## Input Files")
    for item in input_meta:
        name = str(item.get("name", "unknown.pdf"))
        size_mb = item.get("size_mb", 0)
        lines.append(f"- {name} ({size_mb} MB)")
    lines.append("")
    lines.append("## Extracted Text")
    lines.append(
        "No readable text could be extracted locally. For scanned/image PDFs, enable open-source OCR (OCRmyPDF/Tesseract) or configure LIGHTON_OCR_ENDPOINT."
    )
    return "\n".join(lines)


def normalize_extracted_text(raw_text: str) -> str:
    normalized = raw_text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.strip() for line in normalized.split("\n")]
    paragraphs: list[str] = []
    current = ""

    def flush() -> None:
        nonlocal current
        if current:
            paragraphs.append(current.strip())
            current = ""

    for line in lines:
        if not line:
            flush()
            continue

        line = _MULTISPACE_RE.sub(" ", line)

        if not current:
            current = line
            continue

        if current.endswith("-") and line and line[0].islower():
            current = current[:-1] + line
            continue

        current_is_list = bool(_BULLET_RE.match(current))
        next_is_list = bool(_BULLET_RE.match(line))
        current_is_heading = len(current) <= 80 and bool(_HEADING_RE.match(current))
        next_is_heading = len(line) <= 80 and bool(_HEADING_RE.match(line))
        current_ends_sentence = bool(_PUNCT_END_RE.search(current))

        if current_is_list or next_is_list or current_is_heading or next_is_heading:
            flush()
            current = line
            continue

        if current_ends_sentence:
            flush()
            current = line
            continue

        current = f"{current} {line}"

    flush()

    deduped: list[str] = []
    prev = None
    for paragraph in paragraphs:
        if not paragraph:
            continue
        if paragraph == prev:
            continue
        deduped.append(paragraph)
        prev = paragraph

    return "\n\n".join(deduped)


def extract_pdf_text_sections(pdf_bytes: bytes) -> list[tuple[int, str]]:
    # Structural extraction first (headings, lists, tables); pypdf is fallback.
    sections = _extract_sections_with_fitz(pdf_bytes)
    if sections:
        return sections
    return _extract_sections_with_pypdf(pdf_bytes)


def _extract_sections_with_pypdf(pdf_bytes: bytes) -> list[tuple[int, str]]:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception as exc:  # pragma: no cover - depends on file
            raise RuntimeError(f"encrypted_pdf:{exc}") from exc

    sections: list[tuple[int, str]] = []
    for page_index, page in enumerate(reader.pages, start=1):
        raw = page.extract_text() or ""
        text = normalize_extracted_text(raw)
        if text:
            sections.append((page_index, text))

    return sections


_LIST_LINE_RE = re.compile(r"^(?:[-–—*•◦▪‣·]+|\d+[.)])(?:\s+|$)")
_MD_HEADING_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")


def _fitz_is_bold(span: dict[str, Any]) -> bool:
    font = str(span.get("font") or "").lower()
    if any(
        token in font
        for token in ("bold", "black", "heavy", "semibold", "demibold", "-bd")
    ):
        return True
    # PyMuPDF: bit 16 == bold.
    return bool(int(span.get("flags") or 0) & 16)


def _fitz_is_italic(span: dict[str, Any]) -> bool:
    font = str(span.get("font") or "").lower()
    if "italic" in font or "oblique" in font:
        return True
    # PyMuPDF: bit 2 == italic.
    return bool(int(span.get("flags") or 0) & 2)


def _fitz_block_metrics(block: dict[str, Any]) -> tuple[float, float]:
    weighted: dict[float, int] = {}
    bold_chars = 0
    total_chars = 0
    for line in block.get("lines", []):
        for span in line.get("spans", []):
            sample = str(span.get("text") or "")
            chars = len(sample.strip())
            if not chars:
                continue
            size = round(float(span.get("size") or 0.0), 1)
            if size <= 0:
                continue
            weighted[size] = weighted.get(size, 0) + chars
            total_chars += chars
            if _fitz_is_bold(span):
                bold_chars += chars
    if not weighted:
        return 0.0, 0.0
    dominant = max(weighted.items(), key=lambda kv: kv[1])[0]
    bold_ratio = (bold_chars / total_chars) if total_chars else 0.0
    return dominant, bold_ratio


def _fitz_render_span(span: dict[str, Any]) -> str:
    text = _MULTISPACE_RE.sub(" ", str(span.get("text") or ""))
    if not text.strip():
        return ""
    leading = text[: len(text) - len(text.lstrip())]
    trailing = text[len(text.rstrip()):]
    core = text.strip()
    bold = _fitz_is_bold(span)
    italic = _fitz_is_italic(span)
    if bold and italic:
        core = f"***{core}***"
    elif bold:
        core = f"**{core}**"
    elif italic:
        core = f"*{core}*"
    return f"{leading}{core}{trailing}"


def _fitz_render_line(line: dict[str, Any]) -> str:
    rendered = "".join(_fitz_render_span(span) for span in line.get("spans", []))
    return _MULTISPACE_RE.sub(" ", rendered).strip()


def _fitz_canonical_list_line(text: str) -> str:
    ordered = re.match(r"^(\d+)[.)]\s+(.*)$", text)
    if ordered:
        return f"{ordered.group(1)}. {ordered.group(2)}"
    bulletless = _LIST_LINE_RE.sub("", text, count=1).strip()
    return f"- {bulletless}" if bulletless else ""


def _fitz_join_lines(lines: list[str]) -> str:
    result = ""
    for line in lines:
        if not line:
            continue
        if not result:
            result = line
        elif (
            result.endswith("-")
            and not line[0].isupper()
            and not _LIST_LINE_RE.match(line)
        ):
            result = result[:-1] + line
        else:
            result = f"{result} {line}"
    return result.strip()


def _fitz_render_block(block: dict[str, Any], body_size: float) -> list[str]:
    raw_lines = [_fitz_render_line(line) for line in block.get("lines", [])]
    lines = [line for line in raw_lines if line]
    if not lines:
        return []

    list_lines = [line for line in lines if _LIST_LINE_RE.match(line)]
    if list_lines and len(list_lines) >= len(lines):
        rendered: list[str] = []
        for line in lines:
            canonical = _fitz_canonical_list_line(line)
            if canonical:
                rendered.append(canonical)
        return rendered

    text = _fitz_join_lines(lines)
    if not text:
        return []

    # Heading detection: size relative to the dominant body size, short
    # content, and no sentence-ending punctuation.
    size, bold_ratio = _fitz_block_metrics(block)
    ratio = (size / body_size) if body_size > 0 else 1.0
    plain = _MD_HEADING_BOLD_RE.sub(r"\1", text)
    short = len(plain) <= 120
    has_alpha = any(ch.isalpha() for ch in plain)
    all_caps = has_alpha and plain.upper() == plain and len(plain) <= 90
    if short and not _PUNCT_END_RE.search(plain):
        if ratio >= 1.45:
            return [f"# {plain}"]
        if ratio >= 1.2:
            return [f"## {plain}"]
        if ratio >= 1.05 and (all_caps or bold_ratio >= 0.75):
            return [f"### {plain}"]
    return [text]


def _fitz_overlap_ratio(
    box: tuple[float, float, float, float],
    other: tuple[float, float, float, float],
) -> float:
    width = max(box[2] - box[0], 0.0)
    height = max(box[3] - box[1], 0.0)
    area = width * height
    if area <= 0:
        return 0.0
    inter_w = min(box[2], other[2]) - max(box[0], other[0])
    inter_h = min(box[3], other[3]) - max(box[1], other[1])
    if inter_w <= 0 or inter_h <= 0:
        return 0.0
    return (inter_w * inter_h) / area


def _fitz_table_markdown(rows: list[list[Any]]) -> str:
    cleaned: list[list[str]] = []
    for row in rows:
        cells = [
            _MULTISPACE_RE.sub(" ", str(cell or "")).strip().replace("|", "\\|")
            for cell in row
        ]
        if any(cells):
            cleaned.append(cells)
    if len(cleaned) < 2:
        return ""
    width = max(len(row) for row in cleaned)
    cleaned = [row + [""] * (width - len(row)) for row in cleaned]
    lines = [
        "| " + " | ".join(cleaned[0]) + " |",
        "| " + " | ".join(["---"] * width) + " |",
    ]
    for row in cleaned[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def _fitz_page_tables(page: Any) -> list[tuple[tuple[float, ...], str]]:
    try:
        finder = page.find_tables()
        candidates = list(getattr(finder, "tables", None) or [])
    except Exception:
        return []
    tables: list[tuple[tuple[float, ...], str]] = []
    for candidate in candidates:
        try:
            rows = candidate.extract()
            bbox = tuple(candidate.bbox)
        except Exception:
            continue
        markdown = _fitz_table_markdown(rows)
        if markdown:
            tables.append((bbox, markdown))
    return tables


def _fitz_page_markdown(page: Any, body_size: float) -> str:
    try:
        data = page.get_text("dict")
    except Exception:
        return ""

    tables = _fitz_page_tables(page)
    table_boxes = [bbox for bbox, _ in tables]

    text_items: list[tuple[tuple[float, ...], str]] = []
    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        bbox = tuple(block.get("bbox") or (0.0, 0.0, 0.0, 0.0))
        if any(_fitz_overlap_ratio(bbox, box) > 0.2 for box in table_boxes):
            # Cell text inside an extracted table is already part of it.
            continue
        rendered = _fitz_render_block(block, body_size)
        if rendered:
            text_items.append((bbox, "\n\n".join(rendered)))

    parts: list[str] = []
    pending_tables = sorted(
        tables, key=lambda item: (item[0][1] + item[0][3]) / 2
    )
    table_index = 0
    for bbox, markdown in text_items:
        center = (bbox[1] + bbox[3]) / 2
        while (
            table_index < len(pending_tables)
            and (pending_tables[table_index][0][1] + pending_tables[table_index][0][3]) / 2
            < center
        ):
            parts.append(pending_tables[table_index][1])
            table_index += 1
        parts.append(markdown)
    while table_index < len(pending_tables):
        parts.append(pending_tables[table_index][1])
        table_index += 1

    return "\n\n".join(parts).strip()


def _fitz_body_size(document: Any) -> float:
    weights: dict[float, int] = {}
    try:
        page_count = int(document.page_count)
    except Exception:
        page_count = 0
    for page_index in range(min(page_count, 3)):
        try:
            data = document[page_index].get_text("dict")
        except Exception:
            continue
        for block in data.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    sample = str(span.get("text") or "").strip()
                    if len(sample) < 3:
                        continue
                    size = round(float(span.get("size") or 0.0), 1)
                    if size <= 0:
                        continue
                    weights[size] = weights.get(size, 0) + len(sample)
    if not weights:
        return 0.0
    return max(weights.items(), key=lambda kv: kv[1])[0]


def _extract_sections_with_fitz(pdf_bytes: bytes) -> list[tuple[int, str]]:
    if fitz is None:
        return []
    try:
        document = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception:
        return []

    sections: list[tuple[int, str]] = []
    try:
        if document.needs_pass:
            try:
                if not document.authenticate(""):
                    return []
            except Exception:
                return []
        body_size = _fitz_body_size(document)
        try:
            page_count = int(document.page_count)
        except Exception:
            page_count = 0
        for page_index in range(page_count):
            try:
                page = document[page_index]
                markdown = _fitz_page_markdown(page, body_size)
            except Exception:
                markdown = ""
            if markdown:
                sections.append((page_index + 1, markdown))
    finally:
        try:
            document.close()
        except Exception:
            pass
    return sections


def _parse_sidecar_sections(sidecar_text: str) -> list[tuple[int, str]]:
    raw = sidecar_text.replace("\r\n", "\n").replace("\r", "\n")
    pages = raw.split("\f")
    sections: list[tuple[int, str]] = []
    for page_index, page_text in enumerate(pages, start=1):
        text = normalize_extracted_text(page_text)
        if text:
            sections.append((page_index, text))
    return sections


def _ocrmypdf_cli_args(
    input_path: str, output_path: str, sidecar_path: str
) -> list[str]:
    args = [ocrmypdf_cmd()]
    args.extend(["--language", ocrmypdf_langs()])
    args.extend(["--jobs", str(ocrmypdf_jobs())])
    args.extend(["--optimize", "0"])
    args.extend(["--output-type", ocrmypdf_output_type()])
    args.extend(["--tesseract-timeout", str(ocrmypdf_tesseract_timeout_sec())])
    args.extend(["--sidecar", sidecar_path])

    if ocrmypdf_force_ocr():
        args.append("--force-ocr")
    else:
        args.append("--skip-text")
    if ocrmypdf_rotate_pages():
        args.append("--rotate-pages")
    if ocrmypdf_deskew():
        args.append("--deskew")
    if ocrmypdf_clean_final():
        args.append("--clean-final")

    args.extend([input_path, output_path])
    return args


def _split_pdf_for_ocrmypdf(pdf_bytes: bytes, temp_dir: str) -> list[str]:
    """Write the PDF to single-page files, or return the whole file as one chunk.

    Multi-page documents are the ones that exceed the free tier's memory, so
    they are processed page by page. A document PyMuPDF cannot open is passed
    through untouched as a single chunk.
    """
    whole_path = os.path.join(temp_dir, "input.pdf")
    if fitz is None:
        return [whole_path]

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception:
        return [whole_path]

    try:
        if doc.page_count <= 1:
            return [whole_path]

        chunks: list[str] = []
        for index in range(doc.page_count):
            single = fitz.open()
            try:
                single.insert_pdf(doc, from_page=index, to_page=index)
                chunk_path = os.path.join(temp_dir, f"chunk-{index:03d}.pdf")
                single.save(chunk_path)
                chunks.append(chunk_path)
            finally:
                single.close()
        return chunks or [whole_path]
    except Exception:
        return [whole_path]
    finally:
        try:
            doc.close()
        except Exception:
            pass


def _append_ocrmypdf_outputs(
    page_pdf: str, page_sidecar: str, output_path: str, sidecar_path: str
) -> bool:
    """Append one OCR'd page PDF and its sidecar text to the merged outputs."""
    if not os.path.exists(page_pdf):
        return False

    if os.path.exists(page_sidecar):
        try:
            with open(page_sidecar, "r", encoding="utf-8", errors="replace") as f:
                text_chunk = f.read()
            if text_chunk.strip():
                mode = "a" if os.path.exists(sidecar_path) else "w"
                with open(sidecar_path, mode, encoding="utf-8") as f:
                    f.write(text_chunk)
                    if not text_chunk.endswith("\n"):
                        f.write("\n")
                    f.write("\f\n")
        except Exception:
            pass

    if fitz is None:
        if os.path.exists(output_path):
            return False
        try:
            with open(page_pdf, "rb") as src, open(output_path, "wb") as dst:
                dst.write(src.read())
        except Exception:
            return False
        return True

    tmp_path = output_path + ".tmp"
    try:
        merged = fitz.open(output_path) if os.path.exists(output_path) else fitz.open()
        addition = fitz.open(page_pdf)
        try:
            merged.insert_pdf(addition)
            merged.save(tmp_path)
        finally:
            merged.close()
            addition.close()
        os.replace(tmp_path, output_path)
    except Exception:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
        return False

    return True


def _run_subprocess_text(
    args: list[str], timeout_sec: float, env: dict[str, str]
) -> subprocess.CompletedProcess[str]:
    started = time.monotonic()
    proc = subprocess.Popen(
        args,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
        start_new_session=True,
    )
    try:
        stdout, stderr = proc.communicate(timeout=timeout_sec)
    except subprocess.TimeoutExpired as exc:
        try:
            if hasattr(os, "killpg"):
                os.killpg(proc.pid, signal.SIGKILL)
            else:  # pragma: no cover - Windows fallback
                proc.kill()
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass
        try:
            stdout, stderr = proc.communicate(timeout=5)
        except Exception:
            stdout, stderr = "", ""
        elapsed = round(time.monotonic() - started, 2)
        raise RuntimeError(f"subprocess_timeout:{int(timeout_sec)}s elapsed={elapsed}s") from exc
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass
        try:
            proc.communicate(timeout=5)
        except Exception:
            pass
        raise

    return subprocess.CompletedProcess(
        args=args,
        returncode=proc.returncode,
        stdout=stdout,
        stderr=stderr,
    )


def extract_pdf_text_sections_with_ocrmypdf(pdf_bytes: bytes) -> list[tuple[int, str]]:
    """OCR a PDF with OCRmyPDF, one page at a time.

    OCRmyPDF rasterises every page in a single pass. On the 512 MB free tier
    that is what pushed multi-page scans over the limit and surfaced as
    502/503, so documents are split into single pages, OCR'd, then merged.
    """
    timeout_sec = ocrmypdf_timeout_sec()

    with tempfile.TemporaryDirectory(prefix="ocrmypdf_") as temp_dir:
        input_path = os.path.join(temp_dir, "input.pdf")
        output_path = os.path.join(temp_dir, "ocr.pdf")
        sidecar_path = os.path.join(temp_dir, "sidecar.txt")

        with open(input_path, "wb") as f:
            f.write(pdf_bytes)

        page_chunks = _split_pdf_for_ocrmypdf(pdf_bytes, temp_dir)
        print(
            "[ocrmypdf:start]",
            {
                "timeout_sec": timeout_sec,
                "lang": ocrmypdf_langs(),
                "jobs": ocrmypdf_jobs(),
                "rotate_pages": ocrmypdf_rotate_pages(),
                "deskew": ocrmypdf_deskew(),
                "force_ocr": ocrmypdf_force_ocr(),
                "pages": len(page_chunks),
            },
        )
        started = time.monotonic()
        try:
            merged_any = False
            for index, chunk_path in enumerate(page_chunks):
                page_output = os.path.join(temp_dir, f"page-{index:03d}.pdf")
                page_sidecar = os.path.join(temp_dir, f"page-{index:03d}.txt")
                args = _ocrmypdf_cli_args(chunk_path, page_output, page_sidecar)
                result = _run_subprocess_text(
                    args=args,
                    timeout_sec=timeout_sec,
                    env=_ocrmypdf_subprocess_env(),
                )
                if result.returncode != 0:
                    stderr = (result.stderr or "").strip()
                    stdout = (result.stdout or "").strip()
                    message = stderr or stdout or f"exit_code={result.returncode}"
                    message = _MULTISPACE_RE.sub(" ", message).strip()
                    if len(message) > 240:
                        message = message[:240].rstrip() + "..."
                    raise RuntimeError(f"ocrmypdf_cli_failed:{message}")

                if _append_ocrmypdf_outputs(
                    page_output, page_sidecar, output_path, sidecar_path
                ):
                    merged_any = True

            if not merged_any:
                raise RuntimeError("ocrmypdf_empty_result")
        except FileNotFoundError as exc:
            raise RuntimeError("ocrmypdf_not_installed") from exc
        except RuntimeError as exc:
            if str(exc).startswith("subprocess_timeout:"):
                raise RuntimeError(f"ocrmypdf_timeout:{int(timeout_sec)}s") from exc
            raise
        except Exception as exc:
            raise RuntimeError(f"ocrmypdf_failed:{exc}") from exc
        finally:
            print("[ocrmypdf:end]", {"elapsed_sec": round(time.monotonic() - started, 2)})

        sections: list[tuple[int, str]] = []
        if os.path.exists(output_path):
            try:
                with open(output_path, "rb") as f:
                    ocr_pdf_bytes = f.read()
                sections = extract_pdf_text_sections(ocr_pdf_bytes)
            except Exception:
                sections = []

        if not sections and os.path.exists(sidecar_path):
            try:
                with open(sidecar_path, "r", encoding="utf-8", errors="replace") as f:
                    sidecar_text = f.read()
                sections = _parse_sidecar_sections(sidecar_text)
            except Exception:
                sections = []

        if not sections:
            raise RuntimeError("ocrmypdf_empty_result")

        return sections


def _otsu_threshold(gray_image: Any) -> int:
    histogram = gray_image.histogram()
    if not histogram or len(histogram) < 256:
        return 180

    counts = histogram[:256]
    total = sum(counts)
    if total <= 0:
        return 180

    weighted_sum = sum(index * count for index, count in enumerate(counts))
    sum_b = 0.0
    weight_b = 0
    best_variance = -1.0
    threshold = 180

    for index, count in enumerate(counts):
        weight_b += count
        if weight_b == 0:
            continue
        weight_f = total - weight_b
        if weight_f == 0:
            break

        sum_b += index * count
        mean_b = sum_b / weight_b
        mean_f = (weighted_sum - sum_b) / weight_f
        variance = weight_b * weight_f * ((mean_b - mean_f) ** 2)

        if variance > best_variance:
            best_variance = variance
            threshold = index

    return max(40, min(threshold, 230))


def _binarize_luma(gray_image: Any) -> Any:
    threshold = _otsu_threshold(gray_image)
    return gray_image.point(lambda px, t=threshold: 255 if px >= t else 0, mode="L")


def _lanczos_resample() -> int:
    if Image is None:
        return 1
    resampling = getattr(Image, "Resampling", None)
    if resampling is not None:
        return int(resampling.LANCZOS)
    return int(getattr(Image, "LANCZOS", 1))


def _build_tesseract_image_variants(image: Any) -> list[tuple[str, Any]]:
    if Image is None:
        return [("raw", image)]

    # The free Render tier gives 512 MB, and a 300 dpi A4 page is 2481x3507 px.
    # Holding every derived variant (plus a 2x upscale) alive at once pushed the
    # process past the limit on multi-page scans, which surfaced as 502/503.
    # Building them one at a time keeps peak memory to the page itself.
    def _variants() -> Any:
        gray = image.convert("L")
        auto = ImageOps.autocontrast(gray) if ImageOps is not None else gray
        yield "gray_auto", auto
        yield "gray", gray

        if ImageFilter is not None:
            denoised = auto.filter(ImageFilter.MedianFilter(size=3))
            yield "gray_auto_median", denoised
        yield "binary_auto", _binarize_luma(auto)
        if ImageFilter is not None:
            yield "binary_auto_median", _binarize_luma(denoised)

        width, height = auto.size
        if max(width, height) < 2600:
            upscaled = auto.resize((width * 2, height * 2), _lanczos_resample())
            yield "gray_auto_2x", upscaled
            yield "binary_auto_2x", _binarize_luma(upscaled)

    deduped: list[tuple[str, Any]] = []
    seen = set()
    for label, variant in _variants():
        key = (label, getattr(variant, "mode", ""), getattr(variant, "size", None))
        if key in seen:
            continue
        seen.add(key)
        deduped.append((label, variant))
    return deduped[: tesseract_max_variants()]


def _ocr_candidate_score(text: str, mean_confidence: float) -> float:
    normalized = normalize_extracted_text(text)
    if not normalized:
        return -1e9

    text_len = len(normalized)
    word_count = len(_WORD_RE.findall(normalized))
    turkish_hits = sum(normalized.count(ch) for ch in _TURKISH_CHARS)
    replacement_hits = normalized.count("\ufffd")
    symbol_noise = normalized.count("|") + normalized.count("~")
    question_marks = normalized.count("?")

    return (
        (mean_confidence * 4.0)
        + float(text_len)
        + float(word_count * 2)
        + float(turkish_hits * 20)
        - float(replacement_hits * 25)
        - float(symbol_noise * 4)
        - float(question_marks * 15)
    )


def _ocr_candidate_good_enough(text: str, mean_confidence: float) -> bool:
    normalized = normalize_extracted_text(text)
    if len(normalized) < 48:
        return False
    if mean_confidence >= 75:
        return True
    turkish_hits = sum(normalized.count(ch) for ch in _TURKISH_CHARS)
    return turkish_hits >= 3 and mean_confidence >= 50


def _looks_like_bad_turkish_ocr(text: str) -> bool:
    normalized = normalize_extracted_text(text)
    if len(normalized) < 24:
        return False
    turkish_hits = sum(normalized.count(ch) for ch in _TURKISH_CHARS)
    question_marks = normalized.count("?")
    return turkish_hits == 0 and question_marks >= 3


def _tesseract_config(psm: str) -> str:
    info = ensure_tesseract_runtime_config()
    parts = [
        f"--oem {tesseract_oem()}",
        f"--psm {psm}",
        "-c preserve_interword_spaces=1",
        f"-c user_defined_dpi={tesseract_dpi()}",
    ]
    tessdata_dir = str(info.get("tessdata_dir") or "").strip()
    if tessdata_dir:
        parts.append(f'--tessdata-dir "{tessdata_dir}"')
    return " ".join(parts)


def _run_tesseract_candidate(image: Any, lang: str, psm: str) -> tuple[str, float]:
    if pytesseract is None:
        return "", -1.0

    config = _tesseract_config(psm)
    timeout_sec = tesseract_call_timeout_sec()
    try:
        text = pytesseract.image_to_string(
            image, lang=lang, config=config, timeout=timeout_sec
        )
    except TypeError:
        text = pytesseract.image_to_string(image, lang=lang, config=config)
    except Exception:
        text = ""

    return text, -1.0


def _run_tesseract_final_passes(
    image: Any, psm: str, timeout_sec: float
) -> str:
    if pytesseract is None:
        return ""

    best_text = ""
    best_score = -1e9

    for lang in tesseract_final_lang_candidates():
        try:
            text = pytesseract.image_to_string(
                image,
                lang=lang,
                config=_tesseract_config(psm),
                timeout=timeout_sec,
            )
        except TypeError:
            text = pytesseract.image_to_string(
                image,
                lang=lang,
                config=_tesseract_config(psm),
            )
        except Exception:
            continue

        score = _ocr_candidate_score(text, 0.0)
        if score > best_score:
            best_score = score
            best_text = text

    return best_text


def extract_pdf_text_sections_with_tesseract(pdf_bytes: bytes) -> list[tuple[int, str]]:
    if fitz is None or pytesseract is None or Image is None:
        raise RuntimeError("open_source_ocr_dependencies_missing")

    dpi = tesseract_dpi()
    zoom = dpi / 72.0
    lang = tesseract_langs()
    psm_candidates = tesseract_psm_candidates()
    sections: list[tuple[int, str]] = []

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as exc:  # pragma: no cover - dependency/file specific
        raise RuntimeError(f"open_source_ocr_pdf_open_failed:{exc}") from exc

    try:
        for page_index in range(doc.page_count):
            page = doc.load_page(page_index)
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            image = Image.open(io.BytesIO(pix.tobytes("png")))
            # Variants hold their own pixel buffers. Without releasing them the
            # 300 dpi pages of a multi-page scan accumulated until the process
            # was killed on the 512 MB free tier, which surfaced as 502/503.
            variants: list[tuple[str, Any]] = []
            try:
                best_text = ""
                best_score = -1e9
                best_variant: Any | None = None
                best_psm = psm_candidates[0] if psm_candidates else "6"

                variants = _build_tesseract_image_variants(image)
                attempts: list[tuple[Any, str]] = []
                if psm_candidates:
                    primary_psm = psm_candidates[0]
                    for _, variant in variants:
                        attempts.append((variant, primary_psm))
                    for extra_psm in psm_candidates[1:]:
                        for _, variant in variants[:2]:
                            attempts.append((variant, extra_psm))

                attempts = attempts[: tesseract_max_attempts()]

                for variant, psm in attempts:
                    candidate_text, candidate_conf = _run_tesseract_candidate(
                        variant, lang, psm
                    )
                    score = _ocr_candidate_score(candidate_text, candidate_conf)
                    if score > best_score:
                        best_score = score
                        best_text = candidate_text
                        best_variant = variant
                        best_psm = psm
                    if _ocr_candidate_good_enough(candidate_text, candidate_conf):
                        break

                raw = best_text
                final_timeout = min(30.0, tesseract_call_timeout_sec() * 2.0)

                def _pick_better_text(current_text: str, candidate_text: str) -> str:
                    if not candidate_text.strip():
                        return current_text
                    if not current_text.strip():
                        return candidate_text
                    current_score = _ocr_candidate_score(current_text, 0.0)
                    candidate_score = _ocr_candidate_score(candidate_text, 0.0)
                    return candidate_text if candidate_score > current_score else current_text

                if best_variant is not None:
                    try:
                        pretty = _run_tesseract_final_passes(
                            best_variant, best_psm, final_timeout
                        )
                        raw = _pick_better_text(raw, pretty)
                    except Exception:
                        pass

                if _looks_like_bad_turkish_ocr(raw):
                    for _, alt_variant in variants:
                        if alt_variant is best_variant:
                            continue
                        try:
                            pretty_alt = _run_tesseract_final_passes(
                                alt_variant,
                                best_psm,
                                final_timeout,
                            )
                            raw = _pick_better_text(raw, pretty_alt)
                            if not _looks_like_bad_turkish_ocr(raw):
                                break
                        except Exception:
                            continue

                if not raw.strip() and attempts:
                    fallback_variant, fallback_psm = attempts[0]
                    try:
                        raw = _run_tesseract_final_passes(
                            fallback_variant,
                            fallback_psm,
                            final_timeout,
                        )
                    except Exception:
                        raw = ""

                if not raw.strip():
                    raise RuntimeError("tesseract_ocr_empty_result")
            except Exception as exc:  # pragma: no cover - tesseract specific
                raise RuntimeError(f"tesseract_ocr_failed:{exc}") from exc
            finally:
                for _, variant in variants:
                    try:
                        if variant is not image:
                            variant.close()
                    except Exception:
                        pass
                try:
                    pix = None
                    image.close()
                except Exception:
                    pass

            text = normalize_extracted_text(raw or "")
            if text:
                sections.append((page_index + 1, text))
    finally:
        doc.close()

    return sections


def download_storage_bytes(sb: Client, bucket: str, path: str) -> bytes:
    result = sb.storage.from_(bucket).download(path)

    if isinstance(result, (bytes, bytearray)):
        return bytes(result)

    if isinstance(result, tuple) and result:
        first = result[0]
        if isinstance(first, (bytes, bytearray)):
            return bytes(first)

    if hasattr(result, "content"):
        return bytes(result.content)

    raise RuntimeError(f"download_failed_unexpected_type:{type(result).__name__}")


def build_markdown_from_extracted_files(file_results: list[dict[str, Any]]) -> str:
    blocks: list[str] = []

    for item in file_results:
        name = str(item.get("name", "unknown.pdf"))
        lines: list[str] = [f"# {name}", ""]

        note = str(item.get("note") or "").strip()
        if note:
            lines.append(f"> {note}")
            lines.append("")

        sections = item.get("sections") or []
        written = False
        if isinstance(sections, list):
            for section in sections:
                if not isinstance(section, (list, tuple)) or len(section) < 2:
                    continue
                text = str(section[1]).strip()
                if not text:
                    continue
                if written:
                    lines.append(_PAGEBREAK_MARKER)
                    lines.append("")
                lines.append(text)
                lines.append("")
                written = True

        if not written:
            lines.append(
                "No readable text could be extracted from this file. "
                "It may be scanned/image-based and needs OCR "
                "(OPEN_SOURCE_OCR_ENABLED=true)."
            )
            lines.append("")

        blocks.append("\n".join(lines).rstrip())

    if not blocks:
        return ""
    return "\n\n".join(blocks) + "\n"


async def call_lighton_ocr_endpoint(
    endpoint_url: str, token: str, input_meta: list[dict[str, Any]]
) -> str:
    payload = {
        "inputs": {
            "prompt": "Extract text from provided document pages.",
            "files": input_meta,
        }
    }
    async with httpx.AsyncClient(timeout=120) as client:
        response = await client.post(
            endpoint_url,
            json=payload,
            headers={"Authorization": f"Bearer {token}"},
        )
        response.raise_for_status()
        data = response.json()
        if isinstance(data, dict):
            if isinstance(data.get("markdown"), str):
                return data["markdown"]
            if isinstance(data.get("text"), str):
                return data["text"]
        return str(data)


_MD_HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
_MD_BULLET_RE = re.compile(r"^[-*•◦▪‣·]\s+(.*)$")
_MD_ORDERED_RE = re.compile(r"^(\d+)[.)]\s+(.*)$")
_MD_PAGEBREAK_RE = re.compile(
    r"^(?:<!--\s*pagebreak\s*-->|\\pagebreak)$", re.IGNORECASE
)
_MD_PAGE_HEADING_RE = re.compile(r"^page\s+\d+$", re.IGNORECASE)
_MD_HR_RE = re.compile(r"^(?:-{3,}|\*{3,}|_{3,})$")
_MD_INLINE_RE = re.compile(
    r"(\*\*\*.+?\*\*\*|\*\*.+?\*\*|(?<!\\)\*[^*\n]+\*|`[^`\n]+`)"
)


def _split_table_row(row: str) -> list[str]:
    stripped = row.strip()
    if stripped.startswith("|"):
        stripped = stripped[1:]
    if stripped.endswith("|"):
        stripped = stripped[:-1]
    return [cell.strip().replace("\\|", "|") for cell in stripped.split("|")]


def _is_table_separator(row: str) -> bool:
    if "|" not in row:
        return False
    cells = _split_table_row(row)
    if not cells:
        return False
    return all(re.fullmatch(r":?-{3,}:?", cell.replace(" ", "")) for cell in cells)


def _style_run(
    run: Any,
    *,
    bold: bool = False,
    italic: bool = False,
    mono: bool = False,
    muted: bool = False,
) -> None:
    if bold:
        run.bold = True
    if italic:
        run.italic = True
    if mono:
        run.font.name = "Consolas"
    if muted:
        run.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)


def _add_inline_runs(paragraph: Any, text: str, **base: bool) -> None:
    position = 0
    for match in _MD_INLINE_RE.finditer(text):
        if match.start() > position:
            _style_run(paragraph.add_run(text[position : match.start()]), **base)
        token = match.group(0)
        if token.startswith("***") and token.endswith("***") and len(token) >= 6:
            inner, bold, italic, mono = token[3:-3], True, True, False
        elif token.startswith("**") and token.endswith("**") and len(token) >= 4:
            inner, bold, italic, mono = token[2:-2], True, False, False
        elif token.startswith("`") and token.endswith("`") and len(token) >= 2:
            inner, bold, italic, mono = token[1:-1], False, False, True
        else:
            inner, bold, italic, mono = token[1:-1], False, True, False
        _style_run(
            paragraph.add_run(inner),
            bold=bool(base.get("bold")) or bold,
            italic=bool(base.get("italic")) or italic,
            mono=bool(base.get("mono")) or mono,
            muted=bool(base.get("muted")),
        )
        position = match.end()
    if position < len(text):
        _style_run(paragraph.add_run(text[position:]), **base)


def _add_docx_table(document: Any, rows: list[list[str]]) -> None:
    if not rows:
        return
    width = max(len(row) for row in rows)
    normalized = [row + [""] * (width - len(row)) for row in rows]
    table = document.add_table(rows=len(normalized), cols=width)
    try:
        table.style = "Table Grid"
    except Exception:  # pragma: no cover - style may be missing
        pass
    for row_index, row_values in enumerate(normalized):
        cells = table.rows[row_index].cells
        for column_index, value in enumerate(row_values):
            if not value:
                continue
            paragraph = cells[column_index].paragraphs[0]
            _add_inline_runs(paragraph, value, bold=row_index == 0)
    # Word expects a paragraph after a table.
    document.add_paragraph("")


def _configure_docx_styles(document: Any) -> None:
    try:
        normal = document.styles["Normal"]
        normal.font.name = "Calibri"
        normal.font.size = Pt(11)
        normal.paragraph_format.space_after = Pt(6)
        normal.paragraph_format.line_spacing = 1.15
    except Exception:  # pragma: no cover - default template dependent
        pass
    for level in range(1, 7):
        try:
            heading_style = document.styles[f"Heading {level}"]
            heading_style.font.name = "Calibri"
            heading_style.font.color.rgb = RGBColor(0x11, 0x18, 0x27)
        except Exception:  # pragma: no cover - default template dependent
            pass


def markdown_to_docx_bytes(markdown: str) -> bytes:
    document = Document()
    _configure_docx_styles(document)

    lines = markdown.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    emitted = False
    index = 0

    while index < len(lines):
        line = lines[index].strip()

        if not line:
            index += 1
            continue

        if (
            line.startswith("|")
            and index + 1 < len(lines)
            and _is_table_separator(lines[index + 1].strip())
        ):
            rows = [_split_table_row(line)]
            index += 2
            while index < len(lines) and lines[index].strip().startswith("|"):
                rows.append(_split_table_row(lines[index].strip()))
                index += 1
            _add_docx_table(document, rows)
            emitted = True
            continue

        if _MD_PAGEBREAK_RE.match(line):
            if emitted:
                document.add_page_break()
            index += 1
            continue

        if _MD_HR_RE.match(line):
            index += 1
            continue

        heading_match = _MD_HEADING_RE.match(line)
        if heading_match:
            level = min(len(heading_match.group(1)), 6)
            content = heading_match.group(2).strip()
            if _MD_PAGE_HEADING_RE.match(content):
                # Legacy "### Page N" headings become real page breaks.
                if emitted:
                    document.add_page_break()
            else:
                paragraph = document.add_paragraph(style=f"Heading {level}")
                _add_inline_runs(paragraph, content)
                emitted = True
            index += 1
            continue

        bullet_match = _MD_BULLET_RE.match(line)
        if bullet_match:
            paragraph = document.add_paragraph(style="List Bullet")
            _add_inline_runs(paragraph, bullet_match.group(1).strip())
            emitted = True
            index += 1
            continue

        ordered_match = _MD_ORDERED_RE.match(line)
        if ordered_match:
            paragraph = document.add_paragraph(style="List Number")
            _add_inline_runs(paragraph, ordered_match.group(2).strip())
            emitted = True
            index += 1
            continue

        if line.startswith(">"):
            paragraph = document.add_paragraph()
            paragraph.paragraph_format.left_indent = Inches(0.4)
            _add_inline_runs(
                paragraph,
                line.lstrip(">").strip(),
                italic=True,
                muted=True,
            )
            emitted = True
            index += 1
            continue

        paragraph = document.add_paragraph()
        _add_inline_runs(paragraph, line)
        emitted = True
        index += 1

    output = io.BytesIO()
    document.save(output)
    return output.getvalue()


def upload_outputs(
    sb: Client, job_id: str, markdown: str, docx_bytes: bytes
) -> tuple[str, str]:
    bucket = os.environ.get("OCR_RESULTS_BUCKET", "ocr-results")
    md_path = f"{job_id}/result.md"
    docx_path = f"{job_id}/result.docx"

    sb.storage.from_(bucket).upload(
        md_path,
        markdown.encode("utf-8"),
        file_options={"content-type": "text/markdown", "upsert": "true"},
    )
    sb.storage.from_(bucket).upload(
        docx_path,
        docx_bytes,
        file_options={
            "content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "upsert": "true",
        },
    )
    return md_path, docx_path


def extract_sections_for_pdf(
    pdf_bytes: bytes, use_oss_ocr: bool | None = None
) -> tuple[list[tuple[int, str]], str]:
    """Extract per-page text sections with optional open-source OCR fallback.

    Returns ``(sections, note)`` where ``note`` describes which OCR path was
    used (empty when embedded text was read directly).
    """
    if use_oss_ocr is None:
        use_oss_ocr = open_source_ocr_enabled()

    sections = extract_pdf_text_sections(pdf_bytes)
    note = ""
    if sections or not use_oss_ocr:
        return sections, note

    ocr_errors: list[str] = []

    if ocrmypdf_enabled():
        try:
            sections = extract_pdf_text_sections_with_ocrmypdf(pdf_bytes)
            if sections:
                note = "Extracted with open-source OCR (OCRmyPDF + Tesseract)."
        except Exception as exc:  # pragma: no cover - env dependent
            ocr_errors.append(f"OCRmyPDF unavailable: {exc}")

    if not sections:
        try:
            sections = extract_pdf_text_sections_with_tesseract(pdf_bytes)
            if sections:
                note = "Extracted with open-source OCR (Tesseract fallback)."
        except Exception as exc:  # pragma: no cover - env dependent
            ocr_errors.append(f"Tesseract OCR unavailable: {exc}")

    if not sections and ocr_errors:
        note = " | ".join(ocr_errors[:2])

    return sections, note


async def process_job(job_id: str) -> dict[str, Any]:
    sb = make_supabase_client()
    row_res = (
        sb.table("ocr_jobs")
        .select("id, input_meta, status")
        .eq("id", job_id)
        .limit(1)
        .execute()
    )
    rows = row_res.data or []
    if not rows:
        raise HTTPException(status_code=404, detail="job_not_found")

    input_meta = rows[0].get("input_meta") or []
    if not isinstance(input_meta, list):
        input_meta = []

    sb.table("ocr_jobs").update(
        {"status": "processing", "progress_pct": 25, "started_at": utc_now()}
    ).eq("id", job_id).execute()

    try:
        resolved_files = coerce_input_files(job_id, input_meta)
        extracted_files: list[dict[str, Any]] = []
        has_local_text = False
        use_oss_ocr = open_source_ocr_enabled()

        for index, file_meta in enumerate(resolved_files):
            name = str(file_meta.get("name", f"input-{index}.pdf"))
            bucket = str(file_meta.get("bucket", "ocr-inputs"))
            path = str(file_meta.get("path", "")).strip()

            progress = 35 + int(((index + 1) / max(len(resolved_files), 1)) * 25)
            sb.table("ocr_jobs").update({"progress_pct": progress}).eq(
                "id", job_id
            ).execute()

            if not path:
                extracted_files.append(
                    {
                        "name": name,
                        "sections": [],
                        "note": "Storage path is missing for this file.",
                    }
                )
                continue

            try:
                pdf_bytes = download_storage_bytes(sb, bucket, path)
                sections, note = extract_sections_for_pdf(
                    pdf_bytes, use_oss_ocr=use_oss_ocr
                )

                has_local_text = has_local_text or bool(sections)
                item: dict[str, Any] = {"name": name, "sections": sections}
                if note:
                    item["note"] = note
                extracted_files.append(item)
            except Exception as exc:  # pragma: no cover - storage/pdf dependent
                extracted_files.append(
                    {
                        "name": name,
                        "sections": [],
                        "note": f"Could not read PDF content: {exc}",
                    }
                )

        endpoint_url = os.environ.get("LIGHTON_OCR_ENDPOINT", "").strip()
        endpoint_token = os.environ.get("LIGHTON_OCR_TOKEN", "").strip()

        if has_local_text:
            markdown = build_markdown_from_extracted_files(extracted_files)
        elif endpoint_url and endpoint_token:
            markdown = await call_lighton_ocr_endpoint(
                endpoint_url, endpoint_token, resolved_files or input_meta
            )
        else:
            markdown = build_markdown_from_extracted_files(extracted_files)
            if not markdown.strip():
                markdown = build_placeholder_markdown(input_meta)

        sb.table("ocr_jobs").update({"progress_pct": 80}).eq("id", job_id).execute()

        docx_bytes = markdown_to_docx_bytes(markdown)
        md_path, docx_path = upload_outputs(sb, job_id, markdown, docx_bytes)

        sb.table("ocr_jobs").update(
            {
                "status": "succeeded",
                "progress_pct": 100,
                "output_md_path": md_path,
                "output_docx_path": docx_path,
                "finished_at": utc_now(),
            }
        ).eq("id", job_id).execute()

        return {
            "job_id": job_id,
            "status": "succeeded",
            "output_md_path": md_path,
            "output_docx_path": docx_path,
        }
    except Exception as exc:
        sb.table("ocr_jobs").update(
            {
                "status": "failed",
                "progress_pct": 100,
                "error_code": "ocr_error",
                "error_message": str(exc),
                "finished_at": utc_now(),
            }
        ).eq("id", job_id).execute()
        raise


@app.get("/healthz")
def healthz() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/internal/debug/ocr")
def debug_ocr_runtime(request: Request) -> dict[str, Any]:
    assert_worker_secret(request)
    return ocr_runtime_debug_info()


class ProcessRequest(BaseModel):
    job_id: str


@app.post("/internal/process")
async def process(request: Request, payload: ProcessRequest) -> dict[str, Any]:
    assert_worker_secret(request)
    return await process_job(payload.job_id)


@app.post("/internal/process/{job_id}")
async def process_with_path(job_id: str, request: Request) -> dict[str, Any]:
    assert_worker_secret(request)
    return await process_job(job_id)


@app.post("/internal/convert")
async def convert_direct(
    request: Request,
    files: list[UploadFile] = File(...),
    job_id: str = Form(""),
    source: str = Form("flutter_app"),
) -> dict[str, Any]:
    """Convert uploaded PDF bytes directly (no storage download needed).

    Persists job + artifacts when Supabase credentials are configured;
    otherwise returns ``docx_base64`` so the client can keep the result.
    """
    assert_worker_secret(request)
    if not files:
        raise HTTPException(status_code=400, detail="files_required")

    payloads: list[tuple[str, str, bytes]] = []
    for index, upload in enumerate(files):
        data = await upload.read()
        if not data:
            continue
        name = (upload.filename or "").strip() or f"input-{index}.pdf"
        payloads.append((name, upload.content_type or "application/pdf", data))
    if not payloads:
        raise HTTPException(status_code=400, detail="empty_files")

    input_meta: list[dict[str, Any]] = [
        {
            "file_id": f"direct-{index}",
            "name": name,
            "mime_type": mime_type,
            "size_mb": round(len(data) / (1024 * 1024), 3),
            "state": "ready",
        }
        for index, (name, mime_type, data) in enumerate(payloads)
    ]
    total_size_mb = round(
        sum(len(data) for _, _, data in payloads) / (1024 * 1024), 3
    )

    try:
        sb = make_supabase_client()
    except Exception:  # pragma: no cover - depends on env
        sb = None

    active_job_id = (job_id or "").strip()
    if sb is not None:
        if active_job_id:
            row = (
                sb.table("ocr_jobs")
                .select("id")
                .eq("id", active_job_id)
                .limit(1)
                .execute()
            )
            if not row.data:
                raise HTTPException(status_code=404, detail="job_not_found")
            patch: dict[str, Any] = {
                "status": "processing",
                "progress_pct": 30,
                "started_at": utc_now(),
                "input_meta": input_meta,
                "total_size_mb": total_size_mb,
            }
            sb.table("ocr_jobs").update(patch).eq("id", active_job_id).execute()
        else:
            insert = (
                sb.table("ocr_jobs")
                .insert(
                    {
                        "source": source,
                        "status": "processing",
                        "progress_pct": 30,
                        "started_at": utc_now(),
                        "input_meta": input_meta,
                        "total_size_mb": total_size_mb,
                    }
                )
                .select("id")
                .single()
                .execute()
            )
            active_job_id = str(insert.data["id"])

    try:
        extracted: list[dict[str, Any]] = []
        for index, (name, _mime, data) in enumerate(payloads):
            sections, note = extract_sections_for_pdf(data)
            item: dict[str, Any] = {"name": name, "sections": sections}
            if note:
                item["note"] = note
            extracted.append(item)
            if sb is not None and active_job_id:
                pct = 40 + int(((index + 1) / len(payloads)) * 35)
                (
                    sb.table("ocr_jobs")
                    .update({"progress_pct": pct})
                    .eq("id", active_job_id)
                    .execute()
                )

        markdown = build_markdown_from_extracted_files(extracted)
        if not markdown.strip():
            markdown = build_placeholder_markdown(input_meta)

        docx_bytes = markdown_to_docx_bytes(markdown)
        result: dict[str, Any] = {
            "job_id": active_job_id or None,
            "status": "succeeded",
            "file_count": len(payloads),
        }

        if sb is not None and active_job_id:
            md_path, docx_path = upload_outputs(
                sb, active_job_id, markdown, docx_bytes
            )
            done = {
                "status": "succeeded",
                "progress_pct": 100,
                "output_md_path": md_path,
                "output_docx_path": docx_path,
                "finished_at": utc_now(),
            }
            sb.table("ocr_jobs").update(done).eq("id", active_job_id).execute()
            result["output_md_path"] = md_path
            result["output_docx_path"] = docx_path
        else:
            result["docx_base64"] = base64.b64encode(docx_bytes).decode("ascii")
            result["markdown"] = markdown

        return result
    except HTTPException:
        raise
    except Exception as exc:
        if sb is not None and active_job_id:
            failed = {
                "status": "failed",
                "progress_pct": 100,
                "error_code": "direct_convert_error",
                "error_message": str(exc),
                "finished_at": utc_now(),
            }
            sb.table("ocr_jobs").update(failed).eq("id", active_job_id).execute()
        raise HTTPException(status_code=500, detail=str(exc)) from exc

