"""Document generation engine for Excel to Word Autofill Automation.

Supports placeholders such as {{Full Name}} in paragraphs, tables (including nested
 tables), and section headers/footers. Text boxes and other drawing-layer objects
are not exposed by python-docx and are outside the supported scope.
"""
from __future__ import annotations
import re
from pathlib import Path
from typing import Any, Iterable
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import load_workbook

PLACEHOLDER_RE = re.compile(r"\{\{\s*(.*?)\s*\}\}")
INVALID_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def normalize_key(value: Any) -> str:
    return str(value or "").strip().casefold()


def excel_value(value: Any) -> str:
    if value is None:
        return ""
    # Keep dates readable without converting them into Excel serial values.
    if hasattr(value, "strftime"):
        try:
            return value.strftime("%Y-%m-%d")
        except Exception:
            pass
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def safe_filename(value: Any, fallback: str = "Record") -> str:
    name = INVALID_FILENAME.sub("_", str(value or "")).strip().rstrip(". ")
    name = re.sub(r"\s+", "_", name)
    if not name or name in {".", ".."}:
        name = fallback
    # Windows reserved device names
    if name.split(".")[0].upper() in {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}:
        name = "_" + name
    return name[:150]


def load_sheet(path: str | Path, sheet_name: str | None = None) -> tuple[list[str], list[dict[str, Any]], list[str]]:
    """Return headers, row dictionaries, and warnings. Uses first row as headers."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name is None:
            sheet_name = wb.sheetnames[0]
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Worksheet not found: {sheet_name}")
        ws = wb[sheet_name]
        values = ws.iter_rows(values_only=True)
        first = next(values, None)
        if not first:
            return [], [], ["The selected worksheet is empty."]
        raw_headers = list(first)
        warnings: list[str] = []
        headers: list[str] = []
        seen: dict[str, int] = {}
        for i, h in enumerate(raw_headers, 1):
            header = str(h).strip() if h is not None else ""
            if not header:
                header = f"Unnamed column {i}"
                warnings.append(f"Column {i} has a blank header; it was named '{header}'.")
            key = normalize_key(header)
            seen[key] = seen.get(key, 0) + 1
            if seen[key] > 1:
                warnings.append(f"Duplicate column heading '{header}' detected. Rename duplicate headers to avoid ambiguity.")
                header = f"{header} ({seen[key]})"
            headers.append(header)
        rows: list[dict[str, Any]] = []
        for row_number, row in enumerate(values, start=2):
            row_values = list(row)
            if not any(v is not None and str(v).strip() != "" for v in row_values):
                continue
            record: dict[str, Any] = {}
            for idx, header in enumerate(headers):
                record[header] = row_values[idx] if idx < len(row_values) else None
            record["__excel_row__"] = row_number
            rows.append(record)
        return headers, rows, warnings
    finally:
        wb.close()


def _iter_paragraphs_in_table(table: Table) -> Iterable[Paragraph]:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                yield p
            for nested in cell.tables:
                yield from _iter_paragraphs_in_table(nested)


def iter_document_paragraphs(doc) -> Iterable[Paragraph]:
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        yield from _iter_paragraphs_in_table(table)
    for section in doc.sections:
        for container in (section.header, section.footer, section.first_page_header,
                          section.first_page_footer, section.even_page_header,
                          section.even_page_footer):
            for p in container.paragraphs:
                yield p
            for table in container.tables:
                yield from _iter_paragraphs_in_table(table)


def detect_placeholders(docx_path: str | Path) -> list[str]:
    doc = Document(str(docx_path))
    found: dict[str, None] = {}
    for paragraph in iter_document_paragraphs(doc):
        for match in PLACEHOLDER_RE.finditer(paragraph.text):
            key = match.group(1).strip()
            if key:
                found.setdefault(key, None)
    return list(found)


def _replace_in_paragraph(paragraph: Paragraph, mapping: dict[str, str]) -> int:
    """Replace placeholders across runs; inserted value takes formatting of first run.

    This handles placeholders split across runs. It preserves runs outside the match,
    but a replacement spanning differently formatted runs necessarily uses the first
    run's formatting for the inserted value.
    """
    runs = paragraph.runs
    if not runs:
        return 0
    # Retry after each replacement so positions remain simple and predictable.
    replacements = 0
    while True:
        text = "".join(run.text for run in runs)
        match = PLACEHOLDER_RE.search(text)
        if not match:
            break
        name = match.group(1).strip()
        value = mapping.get(normalize_key(name))
        if value is None:
            # Leave unmatched placeholders untouched; search for a later matched one.
            start_search = match.end()
            later = PLACEHOLDER_RE.search(text, start_search)
            if not later:
                break
            # Temporarily look for a matched occurrence after current unmatched one.
            matched = None
            for candidate in PLACEHOLDER_RE.finditer(text, start_search):
                if normalize_key(candidate.group(1).strip()) in mapping:
                    matched = candidate
                    break
            if matched is None:
                break
            match = matched
            name = match.group(1).strip()
            value = mapping.get(normalize_key(name))
        start, end = match.span()
        offsets = []
        pos = 0
        for run in runs:
            offsets.append((pos, pos + len(run.text)))
            pos += len(run.text)
        start_idx = next((i for i, (a, b) in enumerate(offsets) if a <= start < b), None)
        if start_idx is None:
            start_idx = next((i for i, (a, b) in enumerate(offsets) if a <= start <= b), 0)
        end_idx = next((i for i, (a, b) in enumerate(offsets) if a < end <= b), None)
        if end_idx is None:
            end_idx = len(runs) - 1
        start_run = runs[start_idx]
        end_run = runs[end_idx]
        start_local = start - offsets[start_idx][0]
        end_local = end - offsets[end_idx][0]
        if start_idx == end_idx:
            start_run.text = start_run.text[:start_local] + str(value) + start_run.text[end_local:]
        else:
            prefix = start_run.text[:start_local]
            suffix = end_run.text[end_local:]
            start_run.text = prefix + str(value)
            for i in range(start_idx + 1, end_idx):
                runs[i].text = ""
            end_run.text = suffix
        replacements += 1
    return replacements


def replace_placeholders(docx_path: str | Path, output_path: str | Path,
                         values_by_column: dict[str, Any],
                         placeholder_to_column: dict[str, str]) -> dict[str, Any]:
    """Generate one document. Mapping keys are placeholder names; values are Excel headers."""
    doc = Document(str(docx_path))
    normalized_values = {normalize_key(k): excel_value(v) for k, v in values_by_column.items()}
    mapping: dict[str, str] = {}
    for placeholder, column in placeholder_to_column.items():
        col_key = normalize_key(column)
        if col_key in normalized_values:
            mapping[normalize_key(placeholder)] = normalized_values[col_key]
    total = 0
    for paragraph in iter_document_paragraphs(doc):
        total += _replace_in_paragraph(paragraph, mapping)
    # Do not save on top of the source template; caller provides output path.
    doc.save(str(output_path))
    leftovers = detect_placeholders(output_path)
    return {"replacements": total, "unresolved_placeholders": leftovers}


def unique_output_path(folder: str | Path, base_name: str) -> Path:
    folder = Path(folder)
    candidate = folder / f"{base_name}.docx"
    if not candidate.exists():
        return candidate
    n = 2
    while True:
        candidate = folder / f"{base_name}_{n}.docx"
        if not candidate.exists():
            return candidate
        n += 1
