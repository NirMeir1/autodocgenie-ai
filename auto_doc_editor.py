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
# Matches any sequence of four or more underscores
RE_PLACEHOLDER = re.compile(r"_{4,}")
REQUIRED_HEADERS = [
    "שם העסק שם",
    "עיר",
    "כתובת",
    "סכום",
    "שנים",
]

def _find_header_indexes(ws: openpyxl.worksheet.worksheet.Worksheet) -> tuple[list[int], Iterable[tuple]]:
    """Return column indexes for ``REQUIRED_HEADERS`` and remaining rows.

    This scans rows until all headers are found (ignoring leading/trailing
    whitespace) and returns the indexes along with an iterator positioned at
    the first data row.
    """

    rows = ws.iter_rows(values_only=True)
    for _row_idx, row in enumerate(rows, start=1):
        cells = [str(c).strip() if c is not None else "" for c in row]
        try:
            idxs = [cells.index(h) for h in REQUIRED_HEADERS]
        except ValueError:
            continue
        return idxs, rows
    raise KeyError(f"Missing columns: {', '.join(REQUIRED_HEADERS)}")

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
    """Replace sequences of underscores in *doc* with ``replacements``.

    Any run of four or more underscores is considered a placeholder and
    replaced sequentially with the next value.
    """

    rep_iter = iter(replacements)

    def _sub(text: str) -> str:
        return RE_PLACEHOLDER.sub(lambda _: _format_value(next(rep_iter, PLACEHOLDER)), text)

    for text_elem in _iter_text_elements(doc):
        if RE_PLACEHOLDER.search(text_elem.text):
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

    indexes, rows = _find_header_indexes(ws)

    with ThreadPoolExecutor(max_workers=workers) as executor:
        for row in rows:
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
