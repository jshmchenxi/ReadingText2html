#!/usr/bin/env python3
"""Export a reading-courseware content.json to an editable .docx file.

Usage:
  python3 scripts/export_lesson_docx.py content.json ./dist
"""

import html
import json
import re
import sys
import zipfile
from pathlib import Path

from export_worksheet_docx import (
    blank_line_runs,
    heading,
    page_break,
    para,
    print_directions,
)


NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def plain(html_text):
    """Strip HTML tags and unescape entities for a print document."""
    if not html_text:
        return ""
    return html.unescape(re.sub(r"<[^>]+>", "", str(html_text))).strip()


def table(headers, rows, widths):
    total_width = sum(widths)
    table_props = (
        '<w:tblPr>'
        f'<w:tblW w:w="{total_width}" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar>'
        '<w:top w:w="40" w:type="dxa"/>'
        '<w:left w:w="80" w:type="dxa"/>'
        '<w:bottom w:w="40" w:type="dxa"/>'
        '<w:right w:w="80" w:type="dxa"/>'
        "</w:tblCellMar></w:tblPr>"
    )
    grid = "<w:tblGrid>" + "".join(
        f'<w:gridCol w:w="{width}"/>' for width in widths
    ) + "</w:tblGrid>"
    body = []

    def make_cell(value, width, bold=False):
        cell_props = (
            f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>'
            '<w:vAlign w:val="center"/></w:tcPr>'
        )
        return cell_props + para(value, spacing_after=40, bold=bold) + "</w:tc>"

    header_cells = "".join(
        make_cell(value, width, bold=True)
        for value, width in zip(headers, widths)
    )
    body.append(f"<w:tr>{header_cells}</w:tr>")
    for row in rows:
        cells = "".join(make_cell(value, width) for value, width in zip(row, widths))
        body.append(f"<w:tr>{cells}</w:tr>")
    return f"<w:tbl>{table_props}{grid}{''.join(body)}</w:tbl>"


def bullet(text, indent=360, bold_prefix=None):
    if bold_prefix:
        return para(
            f"• {bold_prefix}: {text}",
            spacing_after=80,
            indent=indent,
        )
    return para(f"• {text}", spacing_after=80, indent=indent)


def load_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def build_document(data):
    parts = []
    meta = data.get("meta", {})
    tips = data.get("teacherTips", {})
    pre = data.get("preReading", {})
    wr = data.get("whileReading", {})
    post = data.get("postReading", {})

    # ---------- Header ----------
    parts.append(heading("Reading & Writing Lesson", level=0))
    parts.append(heading(meta.get("titleEn", "Reading Lesson"), level=1))
    objectives = meta.get("objectives", [])
    if objectives:
        parts.append(para("Learning objectives:", bold=True, spacing_after=60))
        for objective in objectives:
            parts.append(bullet(objective))
    if meta.get("timeEstimate"):
        parts.append(para(meta["timeEstimate"], spacing_before=60, spacing_after=140))
    if tips:
        parts.append(para("Teacher tips:", bold=True, spacing_before=80,
                          spacing_after=60))
        for key in ("leadin", "debate", "exit"):
            if tips.get(key):
                parts.append(bullet(tips[key], bold_prefix=key.capitalize()))

    # ---------- Pre-Reading ----------
    parts.append(page_break())
    parts.append(heading("Stage 1 · Pre-Reading", level=1))

    parts.append(heading("Lead-in", level=2))
    lead_in = pre.get("leadIn", {})
    if lead_in.get("prompt"):
        parts.append(para(lead_in["prompt"], spacing_after=80))
    if lead_in.get("placeholder"):
        parts.append(para(f"(Prompt space: {lead_in['placeholder']})", italic=True,
                          spacing_after=120))
    parts.append(blank_line_runs(3))

    parts.append(heading("Prediction", level=2))
    prediction = pre.get("prediction", {})
    if prediction.get("hint"):
        parts.append(para(f"Hint: {prediction['hint']}", italic=True, spacing_after=60))
    if prediction.get("prompt"):
        parts.append(para(prediction["prompt"], spacing_after=80))
    parts.append(blank_line_runs(4))

    key_words = pre.get("keyWords", [])
    if key_words:
        parts.append(heading("Blocking Words", level=2))
        rows = [
            [item.get("word", ""), plain(item.get("meaning", "")),
             plain(item.get("example", ""))]
            for item in key_words
        ]
        parts.append(table(["Word", "Meaning", "Example"], rows,
                           [2200, 3300, 3900]))
        parts.append(page_break())

    # ---------- While-Reading ----------
    parts.append(heading("Stage 2 · While-Reading", level=1))

    paragraphs = wr.get("passageParagraphs", [])
    if paragraphs:
        parts.append(heading("Reading Passage", level=2))
        for i, paragraph_html in enumerate(paragraphs, start=1):
            parts.append(para(f"¶{i}  {plain(paragraph_html)}", spacing_after=120))
        parts.append(page_break())

    gist = wr.get("gist", {})
    if gist:
        parts.append(heading("Gist · Main Idea", level=2))
        if gist.get("hint"):
            parts.append(para(f"Hint: {gist['hint']}", italic=True, spacing_after=60))
        if gist.get("question"):
            parts.append(para(gist["question"], bold=True, spacing_after=80))
        options = gist.get("options", [])
        for i, option in enumerate(options):
            parts.append(para(f"{chr(65 + i)}. {option}", spacing_after=40))
        if isinstance(gist.get("correct"), int):
            letter = chr(65 + gist["correct"])
            parts.append(para(f"Answer: {letter}", bold=True, spacing_after=120))
        functions = gist.get("paragraphFunctions", [])
        if functions:
            parts.append(para("Section functions:", bold=True, spacing_after=60))
            for item in functions:
                parts.append(bullet(f"{item.get('p', '')}  {item.get('label', '')}"))
        if gist.get("doneMessage"):
            parts.append(para(plain(gist["doneMessage"]), italic=True,
                              spacing_after=120))

    structure = wr.get("structure", {})
    if structure:
        parts.append(heading("Text Structure", level=2))
        if structure.get("question"):
            parts.append(para(structure["question"], spacing_after=80))
        parts.append(
            para(
                "Layout: " + structure.get("layout", "") + " · Center label: "
                + " / ".join(structure.get("centerLines", []))
                + " · " + structure.get("centerSub", ""),
                spacing_after=80,
            )
        )
        for node in structure.get("nodes", []):
            parts.append(bullet(
                f"{node.get('label', '')} — {node.get('sub', '')}",
                bold_prefix=node.get("tone", ""),
            ))
        if structure.get("summaryHtml"):
            parts.append(para(plain(structure["summaryHtml"]), italic=True,
                              spacing_after=120))

    walkthrough = wr.get("textWalkthrough", [])
    if walkthrough:
        parts.append(heading("Text Walkthrough", level=2))
        for i, card in enumerate(walkthrough, start=1):
            title = card.get("title", f"Section {i}")
            subtitle = card.get("subtitle", "")
            tag = card.get("tag", "")
            label = f"{i}. {title}"
            if subtitle:
                label += f" · {subtitle}"
            if tag:
                label += f"  ({tag})"
            parts.append(heading(label, level=2))

            logic = card.get("paragraphLogic", {})
            steps = logic.get("steps", [])
            if steps:
                chain = " → ".join(
                    f"{step.get('id', '')}. {step.get('labelEn', '')}"
                    f" ({step.get('labelZh', '')})"
                    for step in steps
                )
                parts.append(para("Paragraph logic: " + chain, bold=True,
                                  spacing_after=60))
            devices = logic.get("devices", [])
            if devices:
                device_text = "; ".join(
                    f"{item.get('labelEn', '')} ({item.get('labelZh', '')})"
                    for item in devices
                )
                parts.append(para("Signposts: " + device_text, italic=True,
                                  spacing_after=80))

            questions = card.get("questions", [])
            for j, question in enumerate(questions, start=1):
                q_type = question.get("type", "")
                prompt = question.get("question", "")
                parts.append(para(f"Q{j} [{q_type}] {prompt}", bold=True,
                                  spacing_before=80, spacing_after=40,
                                  keep_with_next=True))
                if question.get("questionZh"):
                    parts.append(para("中文： " + question["questionZh"],
                                      italic=True, spacing_after=40))
                bullets_en = question.get("bulletsEn") or []
                if question.get("answerEn"):
                    parts.append(para("Answer: " + question["answerEn"],
                                      spacing_after=40, indent=360))
                elif bullets_en:
                    for bullet_item in bullets_en:
                        label, explanation = (
                            bullet_item if isinstance(bullet_item, list)
                            else (bullet_item, "")
                        )
                        parts.append(bullet(
                            (label + ": " + explanation).strip(" :"),
                            indent=720,
                        ))
                if question.get("answerZh"):
                    parts.append(para("参考答案： " + question["answerZh"],
                                      italic=True, spacing_after=80, indent=360))

        parts.append(page_break())

    deep_dive = wr.get("deepDive", {})
    if deep_dive:
        parts.append(heading("Deep Dive", level=2))
        title = deep_dive.get("title", "Deep Dive")
        subtitle = deep_dive.get("subtitle", "")
        parts.append(para(title + (f" · {subtitle}" if subtitle else ""), bold=True,
                          spacing_after=60))
        if deep_dive.get("question"):
            parts.append(para(deep_dive["question"], spacing_after=80))
        en_bullets = deep_dive.get("bulletsEn", [])
        zh_bullets = deep_dive.get("bulletsZh", [])
        for k, bullet_item in enumerate(en_bullets):
            label, explanation = (
                bullet_item if isinstance(bullet_item, list) else (bullet_item, "")
            )
            zh = ""
            if k < len(zh_bullets):
                zh_label, zh_explanation = (
                    zh_bullets[k] if isinstance(zh_bullets[k], list)
                    else (zh_bullets[k], "")
                )
                zh = f"｜{zh_label}: {zh_explanation}"
            parts.append(bullet(f"{label}: {explanation}{zh}"))
        if deep_dive.get("summaryEn"):
            parts.append(para("Summary: " + deep_dive["summaryEn"], italic=True,
                              spacing_after=80))
        if deep_dive.get("summaryZh"):
            parts.append(para("总结： " + deep_dive["summaryZh"], italic=True,
                              spacing_after=120))

    # ---------- Language Focus ----------
    power_words = wr.get("powerWords", [])
    phrases = wr.get("phrases", [])
    structure_groups = wr.get("structures", [])
    if power_words or phrases or structure_groups:
        parts.append(heading("Exploring Language", level=1))
        if power_words:
            parts.append(heading("Power Words", level=2))
            rows = [
                [
                    item.get("word", ""),
                    item.get("pos", ""),
                    item.get("meaningEn", ""),
                    item.get("meaningZh", ""),
                    item.get("trait", ""),
                ]
                for item in power_words
            ]
            parts.append(table(
                ["Word", "POS", "English meaning", "中文", "Trait"],
                rows,
                [1500, 700, 3300, 2600, 1300],
            ))
            parts.append(para("", spacing_after=0))
        if phrases:
            parts.append(heading("Phrases", level=2))
            rows = [
                [
                    item.get("word", ""),
                    item.get("pos", ""),
                    item.get("meaningEn", ""),
                    item.get("meaningZh", ""),
                    item.get("trait", ""),
                ]
                for item in phrases
            ]
            parts.append(table(
                ["Phrase", "POS", "English meaning", "中文", "Trait"],
                rows,
                [2200, 700, 3100, 2200, 1200],
            ))
            parts.append(para("", spacing_after=0))
        if structure_groups:
            parts.append(heading("Structures", level=2))
            for group in structure_groups:
                group_title = group.get("title", "")
                group_sub = group.get("subtitle", "")
                group_tag = group.get("tag", "")
                title_line = group_title
                if group_sub:
                    title_line += f" · {group_sub}"
                if group_tag:
                    title_line += f"  ({group_tag})"
                parts.append(para(title_line, bold=True, spacing_before=80,
                                  spacing_after=40))
                for item in group.get("items", []):
                    parts.append(bullet(item.get("text", ""), indent=540))

    # ---------- Comprehension Check ----------
    true_false = wr.get("trueFalse", [])
    vocab_match = wr.get("vocabMatch", [])
    if true_false or vocab_match:
        parts.append(heading("Comprehension Check", level=1))
        if true_false:
            parts.append(heading("True or False", level=2))
            for item in true_false:
                verdict = "True" if item.get("answer") else "False"
                number = item.get("n", "")
                parts.append(para(
                    f"{number}. {item.get('text', '')}  "
                    f"({verdict})",
                    spacing_after=40,
                ))
                if item.get("reason"):
                    parts.append(para("Reason: " + item["reason"], italic=True,
                                      spacing_after=80, indent=360))
        if vocab_match:
            parts.append(heading("Vocabulary Match", level=2))
            parts.append(para(wr.get("vocabMatchHint", ""), italic=True,
                              spacing_after=80))
            rows = [
                [item.get("word", ""), item.get("meaning", "")]
                for item in vocab_match
            ]
            parts.append(table(["Word", "Meaning"], rows, [2400, 7000]))
            if wr.get("vocabDoneMessage"):
                parts.append(para(plain(wr["vocabDoneMessage"]), italic=True,
                                  spacing_after=80))

    # ---------- Post-Reading ----------
    parts.append(page_break())
    parts.append(heading("Stage 3 · Post-Reading", level=1))

    prediction_check = post.get("predictionCheck", {})
    if prediction_check:
        parts.append(heading("Prediction Check", level=2))
        if prediction_check.get("prompt"):
            parts.append(para(prediction_check["prompt"], spacing_after=80))
        parts.append(blank_line_runs(3))

    speaking = post.get("speaking", {})
    if speaking:
        parts.append(heading("Speaking", level=2))
        if speaking.get("hint"):
            parts.append(para(f"Hint: {speaking['hint']}", italic=True,
                              spacing_after=60))
        for i, prompt_item in enumerate(speaking.get("prompts", []), start=1):
            parts.append(para(f"{i}. {prompt_item.get('prompt', '')}",
                              spacing_after=80))
            parts.append(blank_line_runs(2))

    text_to_self = post.get("textToSelf", {})
    if text_to_self:
        parts.append(heading("Text-to-Self", level=2))
        if text_to_self.get("prompt"):
            parts.append(para(text_to_self["prompt"], spacing_after=80))
        parts.append(blank_line_runs(3))

    text_to_world = post.get("textToWorld", {})
    if text_to_world:
        parts.append(heading("Text-to-World", level=2))
        if text_to_world.get("prompt"):
            parts.append(para(text_to_world["prompt"], spacing_after=80))
        for side in text_to_world.get("sides", []):
            parts.append(bullet(side.get("label", "")))
        if text_to_world.get("extensionPrompt"):
            parts.append(para("Extension: " + text_to_world["extensionPrompt"],
                              spacing_before=80, spacing_after=60))
        parts.append(blank_line_runs(4))

    exit_ticket = post.get("exitTicket", {})
    if exit_ticket:
        parts.append(heading("Exit Ticket", level=2))
        if exit_ticket.get("prompt"):
            parts.append(para(exit_ticket["prompt"], spacing_after=80))
        parts.append(blank_line_runs(3))

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
    stem = data.get("meta", {}).get("filenameStem", "lesson")
    out_path = output_dir / f"{stem}_Reading_Lesson.docx"
    output_dir.mkdir(parents=True, exist_ok=True)
    write_docx(out_path, data)
    print(f"Wrote {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
