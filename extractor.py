import io, re
from dataclasses import dataclass
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

@dataclass
class ExtractOptions:
    include_tables: bool = True
    collapse_blank_runs: bool = True

def _iter_blocks(doc: Document):
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, doc)
        elif child.tag.endswith("}tbl"):
            yield Table(child, doc)

def _table_to_text(tbl: Table) -> str:
    lines = []
    for row in tbl.rows:
        seen = set()
        row_parts = []
        for cell in row.cells:
            tc_id = id(cell._tc)   # dedupe merged cells
            if tc_id in seen:
                continue
            seen.add(tc_id)
            parts = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
            t = "\n".join(parts)
            if t:
                row_parts.append(t)
        if row_parts:
            # Flatten cells vertically (matches the style in your sample)
            lines.append("\n".join(row_parts))
    return "\n".join(lines).strip()

def docx_bytes_to_value_text(data: bytes, opts: ExtractOptions = ExtractOptions()) -> str:
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
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{2,}", "\n", text)

    if opts.collapse_blank_runs:
        text = re.sub(r"\n{4,}", "\n\n\n", text)

    return text.strip()
