"""Create a QA-only HTML rendering from the generated DOCX.

This is a deterministic fallback for hosts where Word/LibreOffice PDF export is
unavailable.  It preserves document order, headings, lists, tables and embedded
images closely enough to inspect clipping, missing content and unreadable assets.
"""

from __future__ import annotations

import html
from pathlib import Path

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph


DOCX = Path(r"D:\PV_MPPT\PV_ESS_MPC_Paper_Reproduction\docs\PV_ESS_MPC_Paper_Reproduction_Manual_Detailed.docx")
OUT = Path(r"D:\PV_MPPT\docs\_qa\pv_ess_mpc_manual_fallback")
ASSETS = OUT / "assets"


def iter_blocks(parent):
    if isinstance(parent, DocumentObject):
        element = parent.element.body
    elif isinstance(parent, _Cell):
        element = parent._tc
    else:
        raise TypeError(type(parent))
    for child in element.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)


def paragraph_image_paths(paragraph: Paragraph, doc: DocumentObject):
    result = []
    for blip in paragraph._p.xpath(".//a:blip"):
        rel_id = blip.get(qn("r:embed"))
        if not rel_id or rel_id not in doc.part.related_parts:
            continue
        part = doc.part.related_parts[rel_id]
        suffix = Path(str(part.partname)).suffix or ".png"
        target = ASSETS / f"{rel_id}{suffix}"
        if not target.exists():
            target.write_bytes(part.blob)
        result.append(target)
    return result


def render_runs(paragraph: Paragraph):
    chunks = []
    for run in paragraph.runs:
        text = html.escape(run.text).replace("\n", "<br>")
        if not text:
            continue
        if run.bold:
            text = f"<strong>{text}</strong>"
        if run.italic:
            text = f"<em>{text}</em>"
        chunks.append(text)
    return "".join(chunks) or html.escape(paragraph.text)


def render_paragraph(paragraph: Paragraph, doc: DocumentObject):
    style = paragraph.style.name if paragraph.style else "Normal"
    escaped = render_runs(paragraph)
    images = paragraph_image_paths(paragraph, doc)
    image_html = "".join(
        f'<img src="assets/{html.escape(path.name)}" alt="embedded figure">' for path in images
    )
    page_break = bool(paragraph._p.xpath('.//w:br[@w:type="page"]'))
    if style.startswith("Heading "):
        try:
            level = max(1, min(3, int(style.split()[-1])))
        except ValueError:
            level = 2
        return f"<h{level}>{escaped}</h{level}>{image_html}"
    if style in {"List Bullet", "List Bullet 2", "List Number", "List Number 2"}:
        klass = "number" if "Number" in style else "bullet"
        return f'<p class="{klass}">{escaped}</p>{image_html}'
    if not escaped and not image_html:
        return '<div class="page-break"></div>' if page_break else "<p>&nbsp;</p>"
    suffix = '<div class="page-break"></div>' if page_break else ""
    return f"<p>{escaped}</p>{image_html}{suffix}"


def render_table(table: Table, doc: DocumentObject):
    header_cells = []
    body_rows = []
    for row_index, row in enumerate(table.rows):
        cells = []
        tag = "th" if row_index == 0 else "td"
        for cell in row.cells:
            inner = []
            for block in iter_blocks(cell):
                if isinstance(block, Paragraph):
                    inner.append(render_paragraph(block, doc))
                else:
                    inner.append(render_table(block, doc))
            cells.append(f"<{tag}>{''.join(inner)}</{tag}>")
        rendered = f"<tr>{''.join(cells)}</tr>"
        if row_index == 0:
            header_cells.append(rendered)
        else:
            body_rows.append(rendered)
    return f"<table><thead>{''.join(header_cells)}</thead><tbody>{''.join(body_rows)}</tbody></table>"


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    ASSETS.mkdir(parents=True, exist_ok=True)
    doc = Document(DOCX)
    body = []
    for block in iter_blocks(doc):
        if isinstance(block, Paragraph):
            body.append(render_paragraph(block, doc))
        else:
            body.append(render_table(block, doc))
    css = r"""
@page { size: Letter; margin: 0.55in 0.65in 0.60in 0.65in; }
* { box-sizing: border-box; }
body { font-family: Calibri, 'Microsoft YaHei', sans-serif; font-size: 9.2pt; line-height: 1.22; color: #1f2937; }
p { margin: 0 0 5pt 0; orphans: 2; widows: 2; }
h1 { color: #1f4e78; font-size: 16pt; margin: 12pt 0 8pt; break-before: page; page-break-before: always; }
h1:first-of-type { break-before: auto; page-break-before: auto; }
h2 { color: #2f75b5; font-size: 12.5pt; margin: 10pt 0 6pt; break-after: avoid; page-break-after: avoid; }
h3 { color: #244062; font-size: 11pt; margin: 8pt 0 5pt; break-after: avoid; page-break-after: avoid; }
.bullet { padding-left: 16pt; }
.bullet::before { content: '• '; margin-left: -12pt; }
.number { padding-left: 18pt; counter-increment: item; }
.number::before { content: counter(item) '. '; margin-left: -16pt; }
table { width: 100%; border-collapse: collapse; table-layout: fixed; margin: 4pt 0 8pt; font-size: 7.2pt; break-inside: auto; }
thead { display: table-header-group; break-after: avoid; page-break-after: avoid; }
tr { break-inside: avoid; page-break-inside: avoid; }
th, td { border: 0.6pt solid #9aa7b2; padding: 3pt 4pt; vertical-align: top; overflow-wrap: anywhere; }
th { background: #d9eaf7; color: #1f4e78; font-weight: 700; }
td p, th p { margin: 0 0 2pt; }
img { display: block; max-width: 100%; max-height: 8.2in; object-fit: contain; margin: 6pt auto 8pt; break-inside: avoid; }
.page-break { break-after: page; page-break-after: always; }
"""
    output = OUT / "manual.html"
    output.write_text(
        "<!doctype html><html><head><meta charset='utf-8'><title>PV ESS MPC manual QA</title>"
        f"<style>{css}</style></head><body>{''.join(body)}</body></html>",
        encoding="utf-8",
    )
    print(output)


if __name__ == "__main__":
    main()
