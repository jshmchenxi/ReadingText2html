#!/usr/bin/env python3
"""Export a reading-courseware worksheet.json to an editable .docx file.

Usage:
  python3 scripts/export_worksheet_docx.py worksheet.json ./dist
"""

import json
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape


NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def font_props(size_half_points=24, bold=False, italic=False, underline=False, color=None):
    parts = []
    rfonts = (
        '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" '
        'w:eastAsia="\u5b8b\u4f53" w:cs="Calibri"/>'
    )
    parts.append(rfonts)
    if bold:
        parts.append("<w:b/>")
    if italic:
        parts.append("<w:i/>")
    if underline:
        parts.append('<w:u w:val="single"/>')
    if color:
        parts.append(f'<w:color w:val="{color}"/>')
    parts.append(f'<w:sz w:val="{size_half_points}"/>')
    parts.append(f'<w:szCs w:val="{size_half_points}"/>')
    return "".join(parts)


def run(text, **style):
    size_half_points = style.pop("size", 24)
    return (
        f'<w:r><w:rPr>{font_props(size_half_points=size_half_points, **style)}</w:rPr>'
        f"<w:t xml:space=\"preserve\">{escape(text)}</w:t></w:r>"
    )


def para(text="", *, spacing_after=120, spacing_before=0, alignment=None, indent=None,
         keep_with_next=False, **style):
    ppr_parts = []
    if keep_with_next:
        ppr_parts.append("<w:keepNext/>")
    ppr_parts.append(f'<w:spacing w:before="{spacing_before}" w:after="{spacing_after}"/>')
    if alignment:
        ppr_parts.append(f'<w:jc w:val="{alignment}"/>')
    if indent is not None:
        ppr_parts.append(f'<w:ind w:left="{indent}"/>')
    ppr = f"<w:pPr>{''.join(ppr_parts)}</w:pPr>"
    content = ppr + (run(text, **style) if text else "")
    return f"<w:p>{content}</w:p>"


def page_break():
    return (
        "<w:p><w:pPr><w:spacing w:after=\"0\"/></w:pPr>"
        "<w:r><w:br w:type=\"page\"/></w:r></w:p>"
    )


def blank_line_runs(count=2):
    paras = []
    for _ in range(count):
        paras.append("<w:p/>")
    return "\n".join(paras)


def heading(text, level=1):
    if level == 0:
        return para(text, size=44, bold=True, spacing_after=160, color="1F4E5F",
                    alignment="center")
    if level == 1:
        return para(text, size=30, bold=True, spacing_before=260, spacing_after=140,
                    color="1F4E5F", keep_with_next=True)
    return para(text, size=26, bold=True, spacing_before=140, spacing_after=80,
                keep_with_next=True)


def print_directions(text):
    """Replace interactive-HTML wording with print/Word-friendly wording."""
    if not text:
        return ""
    cleaned = text.replace(
        "then click Submit to check your objective answers",
        "then check your answers with your teacher or partner",
    )
    cleaned = cleaned.replace(
        "Click a word, then click its meaning.",
        "Write the letter of the correct meaning next to each word.",
    )
    cleaned = cleaned.replace(
        "You may click a word, then click the blank, or drag the word into the blank.",
        "Write the correct word or phrase in each blank.",
    )
    return cleaned


def two_col_table(rows, col1_width=2600, col2_width=6800):
    table_props = (
        '<w:tblPr>'
        '<w:tblW w:w="9400" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar>'
        '<w:top w:w="60" w:type="dxa"/>'
        '<w:left w:w="120" w:type="dxa"/>'
        '<w:bottom w:w="60" w:type="dxa"/>'
        '<w:right w:w="120" w:type="dxa"/>'
        "</w:tblCellMar></w:tblPr>"
    )
    grid = (
        f'<w:tblGrid><w:gridCol w:w="{col1_width}"/>'
        f'<w:gridCol w:w="{col2_width}"/></w:tblGrid>'
    )
    body = []
    for left, right in rows:
        left_cell = (
            f'<w:tc><w:tcPr><w:tcW w:w="{col1_width}" w:type="dxa"/>'
            '<w:vAlign w:val="center"/></w:tcPr>'
            f'{para(left, spacing_after=40)}</w:tc>'
        )
        right_cell = (
            f'<w:tc><w:tcPr><w:tcW w:w="{col2_width}" w:type="dxa"/>'
            '<w:vAlign w:val="center"/></w:tcPr>'
            f'{para(right, spacing_after=40)}</w:tc>'
        )
        body.append(f"<w:tr>{left_cell}{right_cell}</w:tr>")
    return f"<w:tbl>{table_props}{grid}{''.join(body)}</w:tbl>"


def load_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def build_document(data):
    parts = []
    meta = data.get("meta", {})
    matching = data.get("matching", {})
    fill = data.get("fillInBlank", {})
    imitation = data.get("imitation", {})
    para_imit = data.get("paragraphImitation", {})
    summary = data.get("summary", {})

    # Header
    parts.append(heading("Student Worksheet", level=0))
    parts.append(heading(meta.get("titleEn", "Reading Worksheet"), level=1))
    if meta.get("intro"):
        parts.append(para(print_directions(meta["intro"]), italic=True, spacing_after=160))

    # 1. Matching
    parts.append(heading("Part 1 · Matching", level=1))
    parts.append(para(print_directions(matching.get("directions", "")), spacing_after=120))
    items = matching.get("items", [])
    words = [(str(i + 1) + ". " + item["word"]) for i, item in enumerate(items)]
    # A fixed rotation keeps each meaning on a different row from its word.
    offset = 5
    letters = [chr(65 + i) for i in range(len(items))]
    meanings = [
        letters[(i + offset) % len(items)] + ". " + items[(i + offset) % len(items)]["meaning"]
        for i in range(len(items))
    ]
    parts.append(two_col_table(list(zip(words, meanings))))
    parts.append(page_break())

    # 2. Fill in the blanks
    parts.append(heading("Part 2 · Fill in the Blanks", level=1))
    parts.append(para(print_directions(fill.get("directions", "")), spacing_after=120))
    bank = fill.get("wordBank", [])
    if bank:
        parts.append(para("Word Bank:  " + "   |   ".join(bank), bold=True,
                          spacing_after=160))
    for i, item in enumerate(fill.get("items", []), start=1):
        text = (
            f"{str(i)}. {item['before'].rstrip()}  "
            f"________  {item['after'].lstrip()}"
        )
        parts.append(para(text, spacing_after=160))
    parts.append(page_break())

    # 3. Sentence imitation
    parts.append(heading("Part 3 · Sentence Imitation", level=1))
    parts.append(para(imitation.get("directions", ""), spacing_after=160))
    for i, item in enumerate(imitation.get("items", []), start=1):
        parts.append(para(f"Pattern {i}: {item['pattern']}", bold=True,
                          spacing_after=60, keep_with_next=True))
        parts.append(para(f"Example: {item['example']}", italic=True,
                          spacing_after=100))
        scenarios = item.get("scenarios", [])
        for j, scenario in enumerate(scenarios, start=1):
            parts.append(para(f"Scenario {j}: {scenario}", spacing_before=40,
                              spacing_after=40, keep_with_next=True))
            parts.append(blank_line_runs(2))
    parts.append(page_break())

    # 4. Paragraph imitation
    parts.append(heading("Part 4 · Paragraph Imitation", level=1))
    parts.append(para(para_imit.get("directions", ""), spacing_after=120))
    parts.append(para(f"Model paragraph ({para_imit.get('sourceTag', '')}):",
                      bold=True, spacing_after=60, keep_with_next=True))
    parts.append(para(para_imit.get("modelParagraph", ""), italic=True,
                      spacing_after=120, indent=360))
    logic = para_imit.get("logicSteps", [])
    if logic:
        chain = "  →  ".join(str(i + 1) + ". " + step.get("labelEn", "")
                             for i, step in enumerate(logic))
        parts.append(para("Logic chain: " + chain, bold=True, spacing_after=160))
    for i, scenario in enumerate(para_imit.get("scenarios", []), start=1):
        title = scenario.get("titleZh") or scenario.get("titleEn", f"Scenario {i}")
        subtitle = scenario.get("titleEn")
        label = f"{chr(64 + i)}. {title}"
        if subtitle and subtitle != title:
            label += f" ({subtitle})"
        parts.append(para(label, bold=True, spacing_before=80, spacing_after=40,
                          keep_with_next=True))
        parts.append(para(scenario.get("prompt", ""), spacing_after=40))
        parts.append(blank_line_runs(5))
    parts.append(page_break())

    # 5. Summary
    parts.append(heading("Part 5 · Guided Summary", level=1))
    parts.append(para(summary.get("directions", ""), spacing_after=120))
    bank = summary.get("wordBank", [])
    if bank:
        parts.append(para("Word Bank:  " + "   |   ".join(bank), bold=True,
                          spacing_after=160))
    segments = summary.get("segments", [])
    text_parts = []
    for segment in segments:
        if isinstance(segment, str):
            text_parts.append(segment)
        else:
            text_parts.append("________")
    text = "".join(text_parts).replace("\n\n", "\n")
    for chunk in text.split("\n"):
        parts.append(para(chunk.strip(), spacing_after=100))
    parts.append(blank_line_runs(2))

    return "\n".join(parts)


def write_docx(out_path, data):
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{NS_W}"><w:body>'
        f"{build_document(data)}"
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" '
        'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
        "</w:body></w:document>"
    )

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )

    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        "</Relationships>"
    )

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document_xml)


def main(argv):
    if len(argv) not in (2, 3):
        print(__doc__)
        return 2
    source_path = Path(argv[1])
    output_dir = Path(argv[2]) if len(argv) == 3 else source_path.parent
    data = load_json(source_path)
    stem = data.get("meta", {}).get("filenameStem", "worksheet")
    out_path = output_dir / f"{stem}_Student_Worksheet.docx"
    output_dir.mkdir(parents=True, exist_ok=True)
    write_docx(out_path, data)
    print(f"Wrote {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
