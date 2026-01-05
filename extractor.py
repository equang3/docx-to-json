import io
import re
from dataclasses import dataclass
from typing import List, Dict

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


@dataclass
class ExtractOptions:
    include_tables: bool = True
    # In practice this means: collapse multiple newlines down to one
    collapse_blank_runs: bool = True


def _iter_blocks(doc: Document):
    """
    Yields document body blocks in *original order*:
    - Paragraph blocks (w:p)
    - Table blocks (w:tbl)

    python-docx doesn't provide a built-in "iterate paragraphs + tables in order",
    so we iterate the XML children ourselves.
    """
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, doc)
        elif child.tag.endswith("}tbl"):
            yield Table(child, doc)


def _table_to_text(tbl: Table) -> str:
    """
    Converts a docx table into text.

    - Dedupe merged cells by using id(cell._tc)
    - Extract each cell's paragraph texts, joined by "\n"
    - For each row, flatten all non-empty cell texts vertically:
      row becomes one vertical block
    """
    lines = []
    for row in tbl.rows:
        seen = set()
        row_parts = []
        for cell in row.cells:
            tc_id = id(cell._tc)  # dedupe merged cells
            if tc_id in seen:
                continue
            seen.add(tc_id)

            parts = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
            t = "\n".join(parts).strip()
            if t:
                row_parts.append(t)

        if row_parts:
            lines.append("\n".join(row_parts))

    return "\n".join(lines).strip()


def docx_bytes_to_value_text(data: bytes, opts: ExtractOptions = ExtractOptions()) -> str:
    """
    Extracts a single newline-delimited text blob from DOCX bytes.

    Important behavior:
    - Skips empty paragraphs
    - Tables become additional text blocks (if include_tables=True)
    - Normalizes newlines
    - Collapses multiple blank lines if collapse_blank_runs=True
    """
    doc = Document(io.BytesIO(data))
    chunks = []

    for block in _iter_blocks(doc):
        if isinstance(block, Paragraph):
            t = block.text.strip()
            if t:
                chunks.append(t)
        else:
            if opts.include_tables:
                t = _table_to_text(block)
                if t:
                    chunks.append(t)

    text = "\n".join(chunks)

    # Normalize newline styles
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    # Strip trailing spaces before newline
    text = re.sub(r"[ \t]+\n", "\n", text)

    # Optional: collapse blank runs (multiple newlines)
    if opts.collapse_blank_runs:
        text = re.sub(r"\n{2,}", "\n", text)

    return text.strip()


def docx_bytes_to_paras(data: bytes, opts: ExtractOptions = ExtractOptions()) -> List[Dict]:
    """
    Produces the structure you want:
      {"paras": [{"id": 0, "text": "..."}, ...]}

    Rule: "new objects at the new lines" => split extracted text on '\n'.
    """
    value_text = docx_bytes_to_value_text(data, opts=opts)

    # Split on newlines, drop empty/whitespace-only lines
    lines = [ln.strip() for ln in value_text.split("\n")]
    lines = [ln for ln in lines if ln]  # remove empties

    return [{"id": i, "text": ln} for i, ln in enumerate(lines)]
