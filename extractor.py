import io, re
from dataclasses import dataclass
from typing import List
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

_PIPE_ONLY_RE = re.compile(r"^\|+\s*$")

@dataclass
class ExtractOptions:
    include_tables: bool = True
    col_sep: str = " | "          # between table columns
    cell_join: str = " "          # between multiple paragraphs inside one cell

def _iter_blocks(doc: Document):
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, doc)
        elif child.tag.endswith("}tbl"):
            yield Table(child, doc)

def _cell_text(cell, join_with: str) -> str:
    parts = []
    for p in cell.paragraphs:
        t = p.text.strip()
        if t:
            parts.append(t)
    return join_with.join(parts).strip()

def docx_bytes_to_paras(data: bytes, opts: ExtractOptions = ExtractOptions()) -> List[str]:
    doc = Document(io.BytesIO(data))
    out: List[str] = []

    for block in _iter_blocks(doc):
        if isinstance(block, Paragraph):
            t = block.text.strip()
            if t:
                out.append(t)
            continue

        if not opts.include_tables:
            continue

        # Each table row becomes ONE "para"
        for row in block.rows:
            cells = []
            seen = set()

            for cell in row.cells:
                tc_id = id(cell._tc)  # dedupe horizontally merged cells
                if tc_id in seen:
                    continue
                seen.add(tc_id)

                cells.append(_cell_text(cell, opts.cell_join))

            # trim trailing empties
            while cells and not cells[-1]:
                cells.pop()

            if not any(cells):
                continue

            row_text = opts.col_sep.join(cells).strip()

            # drop junk artifact rows like "|" or "|||"
            if not row_text or _PIPE_ONLY_RE.match(row_text):
                continue

            out.append(row_text)

    return out
