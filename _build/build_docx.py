#!/usr/bin/env python3
"""
Build Курсовая_больница.docx from per-section markdown-like text files.

Goals:
- Times New Roman 14, 1.5 line spacing, justified, first-line indent 1.25 cm
  (Russian academic standard, GOST 7.32 / GOST 2.105 style).
- Page size A4, margins: top 2 cm, bottom 2 cm, left 3 cm, right 1 cm.
- Heading 1 (numbered top-level), Heading 2 (numbered subsections).
- Plain Word .docx without external dependencies (just stdlib zipfile + xml).

Source format (one file per section, plain UTF-8 text):
    # 1 Перечень принятых сокращений
    Текст параграфа 1.
    Текст параграфа 2.
    ## 1.1 Подзаголовок (если есть)
    Текст подраздела.
    | col1 | col2 | col3 |        # markdown-like table
    |------|------|------|
    | a    | b    | c    |
    [PAGEBREAK]                      # explicit page break

Lines starting with `#` are H1, `##` are H2, `###` are H3.
Blank lines separate paragraphs. Lines starting with `- ` become bullets.

Output: combined .docx assembled in document order from sections/*.txt
(sorted lexically; use 01_*, 02_*, ... naming).
"""
from __future__ import annotations

import os
import re
import sys
import zipfile
from xml.sax.saxutils import escape

# ----- WordprocessingML helpers ----- #

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def w_p(content: str, style: str | None = None,
        align: str | None = None, indent_first: bool = True,
        spacing_after: int | None = None,
        keep_with_next: bool = False,
        page_break_before: bool = False) -> str:
    """Return a <w:p> with given inline children content."""
    ppr_parts = []
    if style:
        ppr_parts.append(f'<w:pStyle w:val="{style}"/>')
    if keep_with_next:
        ppr_parts.append('<w:keepNext/>')
    if page_break_before:
        ppr_parts.append('<w:pageBreakBefore/>')
    spacing_attrs = ' w:line="360" w:lineRule="auto"'
    if spacing_after is not None:
        spacing_attrs += f' w:after="{spacing_after}"'
    ppr_parts.append(f'<w:spacing{spacing_attrs}/>')
    if align:
        ppr_parts.append(f'<w:jc w:val="{align}"/>')
    if indent_first:
        ppr_parts.append('<w:ind w:firstLine="709"/>')  # 1.25 cm = 709 twips
    ppr = '<w:pPr>' + ''.join(ppr_parts) + '</w:pPr>'
    return f'<w:p>{ppr}{content}</w:p>'


def w_run(text: str, bold: bool = False, italic: bool = False,
          size_pt: int = 14, font: str = 'Times New Roman') -> str:
    rpr = ['<w:rFonts w:ascii="{0}" w:hAnsi="{0}" w:cs="{0}"/>'.format(font)]
    if bold:
        rpr.append('<w:b/><w:bCs/>')
    if italic:
        rpr.append('<w:i/><w:iCs/>')
    rpr.append(f'<w:sz w:val="{size_pt * 2}"/><w:szCs w:val="{size_pt * 2}"/>')
    rpr_xml = '<w:rPr>' + ''.join(rpr) + '</w:rPr>'
    parts = []
    # preserve newlines inside a single inline run as <w:br/>
    chunks = text.split('\n')
    for i, ch in enumerate(chunks):
        if i > 0:
            parts.append('<w:br/>')
        if ch:
            parts.append(f'<w:t xml:space="preserve">{escape(ch)}</w:t>')
    return f'<w:r>{rpr_xml}{"".join(parts)}</w:r>'


# Inline formatting: **bold**, *italic*  → multiple runs
INLINE_RE = re.compile(r'(\*\*[^*]+\*\*|\*[^*]+\*)')


def render_inline(text: str, size_pt: int = 14) -> str:
    out: list[str] = []
    pos = 0
    for m in INLINE_RE.finditer(text):
        if m.start() > pos:
            out.append(w_run(text[pos:m.start()], size_pt=size_pt))
        token = m.group(0)
        if token.startswith('**'):
            out.append(w_run(token[2:-2], bold=True, size_pt=size_pt))
        else:
            out.append(w_run(token[1:-1], italic=True, size_pt=size_pt))
        pos = m.end()
    if pos < len(text):
        out.append(w_run(text[pos:], size_pt=size_pt))
    return ''.join(out)


def w_heading(text: str, level: int) -> str:
    size = {1: 16, 2: 14, 3: 14}[level]
    style = f'Heading{level}'
    run = w_run(text, bold=True, size_pt=size)
    align = 'left'
    return w_p(run, style=style, align=align, indent_first=False,
               spacing_after=240, keep_with_next=True,
               page_break_before=(level == 1))


def w_para(text: str, *, align: str = 'both',
           indent_first: bool = True) -> str:
    return w_p(render_inline(text), align=align,
               indent_first=indent_first)


def w_bullet(text: str) -> str:
    # use a real list — define numbering id 1 in numbering.xml
    ppr = (
        '<w:pPr>'
        '<w:pStyle w:val="ListBullet"/>'
        '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
        '<w:spacing w:line="360" w:lineRule="auto"/>'
        '<w:ind w:left="720" w:hanging="360"/>'
        '<w:jc w:val="both"/>'
        '</w:pPr>'
    )
    return f'<w:p>{ppr}{render_inline(text)}</w:p>'


# ---- table support ---- #

def w_table_row(cells: list[str], *, header: bool = False) -> str:
    tcs = []
    for txt in cells:
        run = render_inline(txt)
        if header:
            run = render_inline(f'**{txt}**')
        ppr = (
            '<w:pPr>'
            '<w:spacing w:line="276" w:lineRule="auto"/>'
            '<w:jc w:val="center"/>' if header else
            '<w:pPr><w:spacing w:line="276" w:lineRule="auto"/>'
            '<w:jc w:val="left"/>'
        )
        # safer: build pPr explicitly
        ppr = (
            '<w:pPr><w:spacing w:line="276" w:lineRule="auto"/>'
            f'<w:jc w:val="{"center" if header else "left"}"/></w:pPr>'
        )
        p = f'<w:p>{ppr}{run}</w:p>'
        # cell properties: borders are at table level
        tcs.append(f'<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>{p}</w:tc>')
    return '<w:tr>' + ''.join(tcs) + '</w:tr>'


def w_table(rows: list[list[str]]) -> str:
    if not rows:
        return ''
    tbl_pr = (
        '<w:tblPr>'
        '<w:tblW w:w="5000" w:type="pct"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tblBorders>'
        '<w:tblLayout w:type="autofit"/>'
        '</w:tblPr>'
    )
    out = ['<w:tbl>', tbl_pr]
    out.append(w_table_row(rows[0], header=True))
    for r in rows[1:]:
        out.append(w_table_row(r))
    out.append('</w:tbl>')
    # Word requires a paragraph after a table
    out.append(w_p('', align='both', indent_first=False))
    return ''.join(out)


# ----- top-level document parts ----- #

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>
'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>
'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      <w:sz w:val="28"/><w:szCs w:val="28"/>
      <w:lang w:val="ru-RU" w:eastAsia="ru-RU" w:bidi="ar-SA"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="0" w:line="360" w:lineRule="auto"/>
      <w:jc w:val="both"/>
      <w:ind w:firstLine="709"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/><w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="240"/>
    <w:ind w:firstLine="0"/><w:jc w:val="left"/></w:pPr>
    <w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="200" w:after="120"/>
    <w:ind w:firstLine="0"/><w:jc w:val="left"/></w:pPr>
    <w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/><w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:spacing w:before="160" w:after="120"/>
    <w:ind w:firstLine="0"/><w:jc w:val="left"/></w:pPr>
    <w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet">
    <w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/>
  </w:style>
</w:styles>
'''

NUMBERING = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="bullet"/>
      <w:lvlText w:val="\u2014"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode"
                     w:uri="http://schemas.microsoft.com/office/word"
                     w:val="15"/>
  </w:compat>
</w:settings>
'''

DOC_HEADER = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:document {W}>'
    '<w:body>'
)
DOC_FOOTER = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1134" w:right="567" w:bottom="1134" w:left="1701"'
    ' w:header="708" w:footer="708" w:gutter="0"/>'
    '</w:sectPr>'
    '</w:body></w:document>'
)


# ----- markdown-lite parser ----- #

def parse_table_block(lines: list[str], i: int) -> tuple[list[list[str]], int]:
    """Parse a markdown-style table starting at lines[i]. Return (rows, next_i)."""
    rows: list[list[str]] = []
    while i < len(lines) and lines[i].lstrip().startswith('|'):
        row = lines[i].strip().strip('|')
        cells = [c.strip() for c in row.split('|')]
        # skip the separator row like |---|---|
        if all(re.match(r'^:?-+:?$', c) for c in cells if c):
            i += 1
            continue
        rows.append(cells)
        i += 1
    return rows, i


def render_section_text(text: str) -> str:
    out: list[str] = []
    lines = text.replace('\r\n', '\n').split('\n')
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if not stripped:
            i += 1
            continue
        if stripped == '[PAGEBREAK]':
            out.append(
                w_p('<w:r><w:br w:type="page"/></w:r>',
                    indent_first=False)
            )
            i += 1
            continue
        if stripped.startswith('### '):
            out.append(w_heading(stripped[4:].strip(), 3))
            i += 1
            continue
        if stripped.startswith('## '):
            out.append(w_heading(stripped[3:].strip(), 2))
            i += 1
            continue
        if stripped.startswith('# '):
            out.append(w_heading(stripped[2:].strip(), 1))
            i += 1
            continue
        if stripped.startswith('- '):
            out.append(w_bullet(stripped[2:]))
            i += 1
            continue
        if stripped.startswith('|'):
            rows, i = parse_table_block(lines, i)
            out.append(w_table(rows))
            continue
        # plain paragraph
        out.append(w_para(stripped))
        i += 1
    return ''.join(out)


def build(out_path: str, sections_dir: str) -> None:
    files = sorted(
        f for f in os.listdir(sections_dir) if f.endswith('.txt')
    )
    body_parts: list[str] = []
    for fname in files:
        with open(os.path.join(sections_dir, fname), encoding='utf-8') as fh:
            body_parts.append(render_section_text(fh.read()))
    document_xml = DOC_HEADER + ''.join(body_parts) + DOC_FOOTER

    out_dir = os.path.dirname(out_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/numbering.xml', NUMBERING)
        z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/document.xml', document_xml)


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print('usage: build_docx.py SECTIONS_DIR OUT.docx', file=sys.stderr)
        sys.exit(2)
    build(sys.argv[2], sys.argv[1])
    print(f'Built {sys.argv[2]} from {sys.argv[1]}')
