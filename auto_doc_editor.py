"""Generate Word documents by filling placeholders using data from an Excel file."""

from __future__ import annotations

import argparse
import io
import re
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Iterable, Any

import openpyxl
from docx import Document
from docx.oxml.ns import qn

OUTPUT_DIR = "AutomaticDocEditor"
PLACEHOLDER = "____"
RE_PLACEHOLDER = re.compile(re.escape(PLACEHOLDER))
REQUIRED_HEADERS = [
    "שם העסק שם",
    "עיר",
    "כתובת",
    "סכום",
    "שנים",
]

def _format_value(value: Any) -> str:
    """Return a string for *value* without trailing ``.0`` for integers."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)

def _iter_text_elements(document: Document):
    """Yield all ``<w:t>`` XML elements in *document* containing text."""

    return document.element.body.iter(qn("w:t"))


def replace_placeholders(doc: Document, replacements: Iterable[Any]) -> Document:
    """Replace ``PLACEHOLDER`` text sequentially in *doc* with *replacements*."""

    rep_iter = iter(replacements)

    def _sub(text: str) -> str:
        return RE_PLACEHOLDER.sub(lambda _: _format_value(next(rep_iter, PLACEHOLDER)), text)

    for text_elem in _iter_text_elements(doc):
        if PLACEHOLDER in text_elem.text:
            text_elem.text = _sub(text_elem.text)
    return doc

def _generate_document(template_bytes: bytes, values: Iterable[Any], output_dir: Path) -> None:
    """Create a document for a single Excel row using *template_bytes*."""

    values = list(values)
    if not values:
        return
    business_name = _format_value(values[0])
    doc = Document(io.BytesIO(template_bytes))
    replace_placeholders(
        doc,
        [_format_value(v) for v in values],
    )
    output_path = output_dir / f"{business_name}.docx"
    doc.save(output_path)
    print(f"Created: {output_path}")


def process_documents(excel_path: str, template_path: str, workers: int | None = None) -> None:
    """Generate Word documents by replacing placeholders based on Excel rows."""

    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(exist_ok=True)

    with open(template_path, "rb") as f:
        template_bytes = f.read()

    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb.active

    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    try:
        indexes = [header_row.index(col) for col in REQUIRED_HEADERS]
    except ValueError as exc:
        missing = [col for col in REQUIRED_HEADERS if col not in header_row]
        raise KeyError(f"Missing columns: {', '.join(missing)}") from exc

    with ThreadPoolExecutor(max_workers=workers) as executor:
        for row in ws.iter_rows(min_row=2, values_only=True):
            values = [row[idx] if idx < len(row) else None for idx in indexes]
            if any(v is not None for v in values):
                executor.submit(
                    _generate_document,
                    template_bytes,
                    values,
                    output_dir,
                )

    wb.close()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Fill Word template using Excel data.")
    parser.add_argument("excel_path", help="Path to the Excel (.xlsx) file")
    parser.add_argument("template_path", help="Path to the Word (.docx) template")
    parser.add_argument(
        "--workers",
        type=int,
        default=None,
        help="Number of parallel workers (default: cpu count)",
    )
    args = parser.parse_args()

    process_documents(args.excel_path, args.template_path, args.workers)
