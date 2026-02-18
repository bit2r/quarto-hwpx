#!/usr/bin/env python3
"""
hwpx_writer.py - Convert Pandoc JSON AST to HWPX document.

Reads Pandoc JSON AST from stdin, generates section0.xml,
updates content.hpf metadata, and repackages Skeleton.hwpx.

Uses raw XML strings (not ElementTree) to avoid namespace/formatting
issues that cause 한글 parse errors.
"""

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
from datetime import datetime
from xml.sax.saxutils import escape as xml_escape

# ── Style mappings ──────────────────────────────────────────────────────────
# (styleIDRef, paraPrIDRef, charPrIDRef)

HEADING_STYLE = {
    1: (2, 2, 7),    # 개요 1 — 22pt bold
    2: (3, 3, 8),    # 개요 2 — 16pt bold
    3: (4, 4, 9),    # 개요 3 — 13pt
    4: (5, 5, 0),    # 개요 4
    5: (6, 6, 0),    # 개요 5
    6: (7, 7, 0),    # 개요 6
}

NORMAL_STYLE = (0, 0, 0)
CAPTION_STYLE = (22, 19, 0)

# charPr entries to inject into header.xml (id, height, bold, fontRef)
HEADING_CHAR_PROPS = [
    (7, 2200, True, 0),   # H1: 22pt, bold, fontRef=0 (per-lang primary)
    (8, 1600, True, 0),   # H2: 16pt, bold, fontRef=0
    (9, 1300, False, 0),  # H3: 13pt, normal, fontRef=0
]

# CodeBlock character properties: D2Coding (fontRef=2), 10pt
CODE_CHAR_PR_ID = 10
CODE_FONT_REF = 2

# Per-language font face mapping (id=0 and id=1 in each fontface block)
LANG_FONT_MAP = {
    'HANGUL':   'NanumSquareOTF',
    'LATIN':    'NimbusSanL',
    'HANJA':    'Noto Sans CJK KR',
    'JAPANESE': 'Noto Sans CJK KR',
    'OTHER':    'NimbusSanL',
    'SYMBOL':   'STIX Two Text',
    'USER':     'NimbusSanL',
}

# Heading spacing: prev (before heading) in HWPUNIT
HEADING_SPACING = {
    2: 800,   # H1 paraPrIDRef=2: 8pt space before
    3: 600,   # H2 paraPrIDRef=3: 6pt space before
    4: 400,   # H3 paraPrIDRef=4: 4pt space before
}

# borderFill id for table cells (solid borders)
TABLE_BORDER_FILL_ID = 3

# ── Paragraph ID generator ─────────────────────────────────────────────────

_para_id_counter = 3121190098

def next_para_id():
    global _para_id_counter
    _para_id_counter += 1
    return str(_para_id_counter)

# ── LaTeX → 한글 수식 스크립트 변환 ──────────────────────────────────────────

def latex_to_hwp_script(latex):
    """Convert LaTeX math to 한글 equation script."""
    s = latex.strip().strip('$')
    # \frac{a}{b} → {a} over {b}
    s = re.sub(r'\\frac\{([^}]*)\}\{([^}]*)\}', r'{\1} over {\2}', s)
    # \sum_{x}^{y} → sum from{x} to{y}
    s = re.sub(r'\\sum_\{([^}]*)\}\^\{([^}]*)\}', r'sum from{\1} to{\2}', s)
    # \int_{x}^{y} → int from{x} to{y}
    s = re.sub(r'\\int_\{([^}]*)\}\^\{([^}]*)\}', r'int from{\1} to{\2}', s)
    # \sqrt{x} → sqrt{x}
    s = re.sub(r'\\sqrt\{([^}]*)\}', r'sqrt{\1}', s)
    # \left( \right) → left( right)
    s = s.replace(r'\left(', 'left(').replace(r'\right)', 'right)')
    # Specific symbol replacements
    for cmd, repl in [('\\geq', '>='), ('\\leq', '<='), ('\\times', 'times'),
                       ('\\cdot', 'cdot'), ('\\infty', 'inf'), ('\\pm', '+-')]:
        s = s.replace(cmd, repl)
    # Remaining \command → command (Greek letters etc.)
    s = re.sub(r'\\([a-zA-Z]+)', r'\1', s)
    return s

# ── Inline extraction ───────────────────────────────────────────────────────

def extract_text(inlines):
    """Extract plain text from Pandoc inline elements."""
    parts = []
    for inline in inlines:
        t = inline.get("t", "")
        c = inline.get("c")
        if t == "Str":
            parts.append(c)
        elif t == "Space":
            parts.append(" ")
        elif t == "SoftBreak":
            parts.append(" ")
        elif t == "LineBreak":
            parts.append("\n")
        elif t in ("Strong", "Emph", "Strikeout", "Superscript",
                    "Subscript", "SmallCaps", "Underline"):
            parts.append(extract_text(c))
        elif t == "Code":
            parts.append(c[1])
        elif t == "Link":
            parts.append(extract_text(c[1]))
        elif t == "Image":
            parts.append(extract_text(c[1]))
        elif t == "Quoted":
            quote_type = c[0].get("t", "DoubleQuote") if isinstance(c[0], dict) else c[0]
            q = '\u201c' if quote_type == "DoubleQuote" else '\u2018'
            q2 = '\u201d' if quote_type == "DoubleQuote" else '\u2019'
            parts.append(q + extract_text(c[1]) + q2)
        elif t == "Cite":
            parts.append(extract_text(c[1]))
        elif t == "Math":
            parts.append(c[1])
        elif t == "Span":
            parts.append(extract_text(c[1]))
    return "".join(parts)


def extract_meta_text(meta_val):
    """Extract plain text from a Pandoc Meta value."""
    if meta_val is None:
        return ""
    t = meta_val.get("t", "")
    c = meta_val.get("c")
    if t == "MetaString":
        return c
    elif t == "MetaInlines":
        return extract_text(c)
    elif t == "MetaList":
        return ", ".join(extract_meta_text(item) for item in c)
    return ""

# ── Layout constants ───────────────────────────────────────────────────────

PAGE_TEXT_WIDTH = 42520   # horzsize: page width minus margins
CHAR_HEIGHT_NORMAL = 1000 # 10pt in HWPUNIT
LINE_SPACING_PCT = 160    # 160% line spacing

# charPr id → height mapping for lineseg calculation
CHAR_HEIGHT_MAP = {
    0: 1000,    # 10pt
    7: 2200,    # H1 22pt
    8: 1600,    # H2 16pt
    9: 1300,    # H3 13pt
    10: 1000,   # CodeBlock 10pt (D2Coding)
}

# ── Lineseg computation ───────────────────────────────────────────────────

def compute_lineseg_xml(text, char_height=CHAR_HEIGHT_NORMAL,
                        line_spacing_pct=LINE_SPACING_PCT,
                        horzsize=PAGE_TEXT_WIDTH):
    """Build <hp:linesegarray> with multi-line entries based on text length."""
    vertsize = char_height
    spacing = int(char_height * (line_spacing_pct - 100) / 100)
    line_height = vertsize + spacing
    baseline = int(char_height * 0.85)

    if not text:
        return (
            f'<hp:linesegarray>'
            f'<hp:lineseg textpos="0" vertpos="0" vertsize="{vertsize}"'
            f' textheight="{vertsize}" baseline="{baseline}"'
            f' spacing="{spacing}" horzpos="0" horzsize="{horzsize}"'
            f' flags="393216"/>'
            f'</hp:linesegarray>'
        )

    line_starts = [0]
    current_width = 0

    for i, ch in enumerate(text):
        if ord(ch) > 0x2000:
            current_width += char_height
        else:
            current_width += char_height // 2

        if current_width > horzsize and (i + 1) < len(text):
            line_starts.append(i + 1)
            current_width = 0

    num_lines = len(line_starts)

    parts = ['<hp:linesegarray>']
    for idx, textpos in enumerate(line_starts):
        vertpos = idx * line_height
        if num_lines == 1:
            flags = 393216   # 0x60000: first + last
        elif idx == 0:
            flags = 131072   # 0x20000: first only
        elif idx == num_lines - 1:
            flags = 262144   # 0x40000: last only
        else:
            flags = 0

        parts.append(
            f'<hp:lineseg textpos="{textpos}" vertpos="{vertpos}"'
            f' vertsize="{vertsize}" textheight="{vertsize}"'
            f' baseline="{baseline}" spacing="{spacing}"'
            f' horzpos="0" horzsize="{horzsize}" flags="{flags}"/>'
        )
    parts.append('</hp:linesegarray>')

    return ''.join(parts)

# ── XML string builders ────────────────────────────────────────────────────

def make_paragraph_xml(text, style_id=0, para_pr_id=0, char_pr_id=0):
    """Build a <hp:p> XML string with computed linesegarray."""
    pid = next_para_id()
    safe_text = xml_escape(text) if text else ""
    char_height = CHAR_HEIGHT_MAP.get(char_pr_id, CHAR_HEIGHT_NORMAL)
    lineseg = compute_lineseg_xml(text, char_height=char_height)
    return (
        f'<hp:p id="{pid}" paraPrIDRef="{para_pr_id}" styleIDRef="{style_id}"'
        f' pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{char_pr_id}"><hp:t>{safe_text}</hp:t></hp:run>'
        f'{lineseg}'
        f'</hp:p>'
    )


def make_equation_paragraph_xml(latex_str):
    """Build a <hp:p> containing an <hp:equation> object."""
    pid = next_para_id()
    script = latex_to_hwp_script(latex_str)
    safe_script = xml_escape(script)
    return (
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="0">'
        f'<hp:equation version="eqEdit" baseLine="0"'
        f' textColor="#000000" baseUnit="1000" lineMode="0" font="">'
        f'<hp:script>{safe_script}</hp:script>'
        f'</hp:equation>'
        f'</hp:run>'
        f'<hp:linesegarray>'
        f'<hp:lineseg textpos="0" vertpos="0" vertsize="1600"'
        f' textheight="1600" baseline="1360" spacing="400"'
        f' horzpos="0" horzsize="42520" flags="393216"/>'
        f'</hp:linesegarray>'
        f'</hp:p>'
    )


def make_table_xml(rows, col_count, caption_text=""):
    """Build table with optional caption, centered, with solid borders."""
    parts = []

    # Caption above table (표 N. caption)
    if caption_text:
        sid, pid, cid = CAPTION_STYLE
        parts.append(make_paragraph_xml(caption_text, sid, pid, cid))

    pid = next_para_id()
    tbl_id = next_para_id()
    page_width = 42520
    col_width = page_width // max(col_count, 1)
    row_height = 1800
    total_height = row_height * len(rows)

    bfid = TABLE_BORDER_FILL_ID

    parts.append(
        f'<hp:p id="{pid}" paraPrIDRef="0" styleIDRef="0"'
        f' pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="0">'
        f'<hp:tbl id="{tbl_id}" zOrder="0" numberingType="TABLE"'
        f' textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0"'
        f' dropcapstyle="None" pageBreak="CELL" repeatHeader="0"'
        f' rowCnt="{len(rows)}" colCnt="{col_count}"'
        f' cellSpacing="0" borderFillIDRef="{bfid}" noAdjust="0">'
        f'<hp:sz width="{page_width}" widthRelTo="ABSOLUTE"'
        f' height="{total_height}" heightRelTo="ABSOLUTE" protect="0"/>'
        f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1"'
        f' allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA"'
        f' horzRelTo="COLUMN" vertAlign="TOP" horzAlign="CENTER"'
        f' vertOffset="0" horzOffset="0"/>'
        f'<hp:outMargin left="0" right="0" top="141" bottom="141"/>'
        f'<hp:inMargin left="0" right="0" top="0" bottom="0"/>'
    )

    for row_idx, row in enumerate(rows):
        parts.append('<hp:tr>')
        for col_idx, cell_text in enumerate(row):
            header = "1" if row_idx == 0 else "0"
            safe_text = xml_escape(cell_text) if cell_text else ""
            cell_pid = next_para_id()
            cell_lineseg = compute_lineseg_xml(cell_text, horzsize=col_width)
            # Header row uses bold charPr (id=9, 13pt) or just center-align
            cell_cpr = "0"
            parts.append(
                f'<hp:tc name="" header="{header}" hasMargin="0"'
                f' protect="0" editable="0" dirty="0" borderFillIDRef="{bfid}">'
                f'<hp:subList id="" textDirection="HORIZONTAL"'
                f' lineWrap="BREAK" vertAlign="CENTER"'
                f' linkListIDRef="0" linkListNextIDRef="0"'
                f' textWidth="0" textHeight="0"'
                f' hasTextRef="0" hasNumRef="0">'
                f'<hp:p id="{cell_pid}" paraPrIDRef="0" styleIDRef="0"'
                f' pageBreak="0" columnBreak="0" merged="0">'
                f'<hp:run charPrIDRef="{cell_cpr}"><hp:t>{safe_text}</hp:t></hp:run>'
                f'{cell_lineseg}'
                f'</hp:p>'
                f'</hp:subList>'
                f'<hp:cellAddr colAddr="{col_idx}" rowAddr="{row_idx}"/>'
                f'<hp:cellSpan colSpan="1" rowSpan="1"/>'
                f'<hp:cellSz width="{col_width}" height="{row_height}"/>'
                f'<hp:cellMargin left="141" right="141" top="141" bottom="141"/>'
                f'</hp:tc>'
            )
        parts.append('</hp:tr>')

    parts.append('</hp:tbl>')
    parts.append('<hp:t></hp:t></hp:run></hp:p>')

    return ''.join(parts)

# ── Block conversion ────────────────────────────────────────────────────────

def convert_blocks(blocks, indent_level=0):
    """Convert Pandoc AST blocks to list of XML strings."""
    xml_parts = []
    indent_prefix = "\u3000" * indent_level  # fullwidth space for indent

    for block in blocks:
        t = block.get("t", "")
        c = block.get("c")

        if t == "Para" or t == "Plain":
            # Standalone math equation → HWPX equation object
            if len(c) == 1 and c[0].get("t") == "Math":
                latex_str = c[0]["c"][1]
                xml_parts.append(make_equation_paragraph_xml(latex_str))
            else:
                text = extract_text(c)
                xml_parts.append(make_paragraph_xml(indent_prefix + text))

        elif t == "Header":
            level = c[0]
            text = extract_text(c[2])
            sid, pid, cid = HEADING_STYLE.get(level, NORMAL_STYLE)
            xml_parts.append(make_paragraph_xml(text, sid, pid, cid))

        elif t == "CodeBlock":
            code_text = c[1]
            for line in code_text.split("\n"):
                xml_parts.append(make_paragraph_xml(
                    indent_prefix + line, char_pr_id=CODE_CHAR_PR_ID))

        elif t == "BulletList":
            for item_blocks in c:
                item_parts = convert_blocks(item_blocks, indent_level)
                if item_parts:
                    item_parts[0] = item_parts[0].replace(
                        '<hp:t>', '<hp:t>• ', 1)
                xml_parts.extend(item_parts)

        elif t == "OrderedList":
            start_num = c[0][0]
            for idx, item_blocks in enumerate(c[1]):
                item_parts = convert_blocks(item_blocks, indent_level)
                if item_parts:
                    num = start_num + idx
                    first = item_parts[0]
                    first = first.replace('<hp:t>', f'<hp:t>{num}. ', 1)
                    item_parts[0] = first
                xml_parts.extend(item_parts)

        elif t == "BlockQuote":
            child_parts = convert_blocks(c, indent_level + 1)
            xml_parts.extend(child_parts)

        elif t == "Table":
            caption = c[1]
            table_head = c[3]
            table_bodies = c[4]

            all_rows = []
            col_count = 0

            head_rows = table_head[1] if len(table_head) > 1 else []
            for row in head_rows:
                cells = row[1] if len(row) > 1 else []
                row_texts = []
                for cell in cells:
                    cell_blocks = cell[4] if len(cell) > 4 else []
                    cell_text = ""
                    for cb in cell_blocks:
                        if cb.get("t") in ("Para", "Plain"):
                            cell_text += extract_text(cb.get("c", []))
                    row_texts.append(cell_text)
                if row_texts:
                    col_count = max(col_count, len(row_texts))
                    all_rows.append(row_texts)

            for tbody in table_bodies:
                body_rows = tbody[3] if len(tbody) > 3 else []
                for row in body_rows:
                    cells = row[1] if len(row) > 1 else []
                    row_texts = []
                    for cell in cells:
                        cell_blocks = cell[4] if len(cell) > 4 else []
                        cell_text = ""
                        for cb in cell_blocks:
                            if cb.get("t") in ("Para", "Plain"):
                                cell_text += extract_text(cb.get("c", []))
                        row_texts.append(cell_text)
                    if row_texts:
                        col_count = max(col_count, len(row_texts))
                        all_rows.append(row_texts)

            for row in all_rows:
                while len(row) < col_count:
                    row.append("")

            # Extract caption text
            cap_text = ""
            if caption and len(caption) > 1 and caption[1]:
                for cb in caption[1]:
                    if cb.get("t") in ("Para", "Plain"):
                        cap_text += extract_text(cb.get("c", []))

            if all_rows and col_count > 0:
                xml_parts.append(make_table_xml(all_rows, col_count, cap_text))

        elif t == "HorizontalRule":
            xml_parts.append(make_paragraph_xml("━" * 30))

        elif t == "Div":
            xml_parts.extend(convert_blocks(c[1], indent_level))

        elif t == "DefinitionList":
            for item in c:
                term_text = extract_text(item[0])
                xml_parts.append(make_paragraph_xml(indent_prefix + term_text))
                for def_blocks in item[1]:
                    xml_parts.extend(convert_blocks(def_blocks, indent_level + 1))

        elif t == "LineBlock":
            for line_inlines in c:
                line_text = extract_text(line_inlines)
                xml_parts.append(make_paragraph_xml(indent_prefix + line_text))

    return xml_parts

# ── Section XML builder ─────────────────────────────────────────────────────

def build_title_block_xml(title="", subtitle="", author="", date_str=""):
    """Build title block paragraphs from document metadata."""
    title_parts = []
    if title:
        # Normal paragraph style (no outline numbering) + H1 charPr (22pt bold)
        title_parts.append(make_paragraph_xml(title, style_id=0, para_pr_id=0, char_pr_id=7))
    if subtitle:
        # Normal paragraph style + H2 charPr (16pt bold)
        title_parts.append(make_paragraph_xml(subtitle, style_id=0, para_pr_id=0, char_pr_id=8))
    if author or date_str:
        meta_line = " | ".join(part for part in [author, date_str] if part)
        title_parts.append(make_paragraph_xml(meta_line))
    if title_parts:
        title_parts.append(make_paragraph_xml(""))  # blank separator
    return title_parts


def build_section_xml(skeleton_path, blocks, title="", subtitle="",
                      author="", date_str=""):
    """Build section0.xml from Skeleton template + AST blocks."""
    with zipfile.ZipFile(skeleton_path, 'r') as z:
        original = z.read('Contents/section0.xml').decode('utf-8')

    sec_tag_start = original.index('<hs:sec')
    sec_tag_end = original.index('>', sec_tag_start) + 1
    xml_header_and_open = original[:sec_tag_end]

    first_p_start = original.index('<hp:p ')
    first_p_end = original.index('</hp:p>') + len('</hp:p>')
    first_paragraph = original[first_p_start:first_p_end]

    title_block = build_title_block_xml(title, subtitle, author, date_str)
    content_parts = convert_blocks(blocks)
    if not title_block and not content_parts:
        content_parts = [make_paragraph_xml("")]

    parts = [
        xml_header_and_open,
        first_paragraph,
        ''.join(title_block),
        ''.join(content_parts),
        '</hs:sec>'
    ]

    return ''.join(parts)

# ── Header.xml updater ─────────────────────────────────────────────────────

def _make_charpr_xml(cpr_id, height, bold=False, font_ref=0):
    """Build a <hh:charPr> XML string for heading character properties."""
    bold_attr = ' bold="1"' if bold else ''
    return (
        f'<hh:charPr id="{cpr_id}" height="{height}"'
        f' textColor="#000000" shadeColor="none"'
        f' useFontSpace="0" useKerning="0" symMark="NONE"'
        f' borderFillIDRef="2"{bold_attr}>'
        f'<hh:fontRef hangul="{font_ref}" latin="{font_ref}"'
        f' hanja="{font_ref}" japanese="{font_ref}"'
        f' other="{font_ref}" symbol="{font_ref}" user="{font_ref}"/>'
        f'<hh:ratio hangul="100" latin="100" hanja="100"'
        f' japanese="100" other="100" symbol="100" user="100"/>'
        f'<hh:spacing hangul="0" latin="0" hanja="0"'
        f' japanese="0" other="0" symbol="0" user="0"/>'
        f'<hh:relSz hangul="100" latin="100" hanja="100"'
        f' japanese="100" other="100" symbol="100" user="100"/>'
        f'<hh:offset hangul="0" latin="0" hanja="0"'
        f' japanese="0" other="0" symbol="0" user="0"/>'
        f'<hh:underline type="NONE" shape="SOLID" color="#000000"/>'
        f'<hh:strikeout shape="NONE" color="#000000"/>'
        f'<hh:outline type="NONE"/>'
        f'<hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/>'
        f'</hh:charPr>'
    )


def _make_table_borderfill_xml():
    """Build a <hh:borderFill> with solid borders for table cells."""
    return (
        f'<hh:borderFill id="{TABLE_BORDER_FILL_ID}" threeD="0" shadow="0"'
        f' centerLine="NONE" breakCellSeparateLine="0">'
        f'<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
        f'<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
        f'<hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/>'
        f'<hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/>'
        f'<hh:topBorder type="SOLID" width="0.12 mm" color="#000000"/>'
        f'<hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/>'
        f'<hh:diagonal type="NONE" width="0.12 mm" color="#000000"/>'
        f'</hh:borderFill>'
    )


def _make_font_xml(font_id, face_name):
    """Build a single <hh:font> entry."""
    return (
        f'<hh:font id="{font_id}" face="{face_name}" type="TTF" isEmbedded="0">'
        f'<hh:typeInfo familyType="FCAT_GOTHIC" weight="6" proportion="4"'
        f' contrast="0" strokeVariation="1" armStyle="1" letterform="1"'
        f' midline="1" xHeight="1"/>'
        f'</hh:font>'
    )


def _replace_fontface_block(match):
    """Replace font entries within a single <hh:fontface> block per language."""
    lang = match.group(1)
    primary_font = LANG_FONT_MAP.get(lang, 'NimbusSanL')
    return (
        f'<hh:fontface lang="{lang}" fontCnt="3">'
        + _make_font_xml(0, primary_font)
        + _make_font_xml(1, primary_font)
        + _make_font_xml(2, 'D2Coding')
        + '</hh:fontface>'
    )


def update_header_xml(header_xml):
    """Inject heading charPr, table borderFill, heading spacing, and font changes."""

    # 0. Replace each fontface block with per-language font mapping
    header_xml = re.sub(
        r'<hh:fontface lang="(\w+)"[^>]*>.*?</hh:fontface>',
        _replace_fontface_block,
        header_xml,
        flags=re.DOTALL
    )

    # 1. Add charPr entries for heading fonts + code block
    new_charpr = ''.join(
        _make_charpr_xml(cpr_id, height, bold, font_ref)
        for cpr_id, height, bold, font_ref in HEADING_CHAR_PROPS
    )
    # Add CodeBlock charPr (id=10, 10pt, not bold, fontRef=2 for D2Coding)
    new_charpr += _make_charpr_xml(CODE_CHAR_PR_ID, 1000, False, CODE_FONT_REF)

    marker = '</hh:charProperties>'
    header_xml = header_xml.replace(marker, new_charpr + marker)
    header_xml = re.sub(
        r'(<hh:charProperties\s+itemCnt=")7(")',
        r'\g<1>11\2', header_xml
    )

    # 2. Add borderFill for table cells (solid borders)
    bf_marker = '</hh:borderFillList>'
    header_xml = header_xml.replace(
        bf_marker, _make_table_borderfill_xml() + bf_marker
    )
    header_xml = re.sub(
        r'(<hh:borderFillList\s+itemCnt=")2(")',
        r'\g<1>3\2', header_xml
    )

    # 3. Add spacing before headings (modify paraPr prev values)
    for para_pr_id, prev_val in HEADING_SPACING.items():
        # Find paraPr with this id and replace prev value=0 → prev_val
        # Pattern: within paraPr id="N", replace <hc:prev value="0"
        pattern = (
            rf'(<hh:paraPr\s+id="{para_pr_id}"[^>]*>.*?)'
            rf'(<hc:prev\s+value=")0(")'
        )
        header_xml = re.sub(pattern, rf'\g<1>\g<2>{prev_val}\3',
                            header_xml, flags=re.DOTALL)

    return header_xml

# ── Content.hpf updater ────────────────────────────────────────────────────

def update_content_hpf(hpf_xml, title="", author="", date_str=""):
    """Update metadata in content.hpf using string replacement."""
    now = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")

    if title:
        hpf_xml = re.sub(
            r'(<opf:title)(?:/>|>(.*?)</opf:title>)',
            rf'\1>{xml_escape(title)}</opf:title>',
            hpf_xml)

    if author:
        safe_author = xml_escape(author)
        hpf_xml = re.sub(
            r'(<opf:meta name="creator" content="text")(?:/>|>(.*?)</opf:meta>)',
            rf'\1>{safe_author}</opf:meta>',
            hpf_xml)
        hpf_xml = re.sub(
            r'(<opf:meta name="lastsaveby" content="text")(?:/>|>(.*?)</opf:meta>)',
            rf'\1>{safe_author}</opf:meta>',
            hpf_xml)

    hpf_xml = re.sub(
        r'(<opf:meta name="ModifiedDate" content="text")(?:/>|>(.*?)</opf:meta>)',
        rf'\1>{now}</opf:meta>',
        hpf_xml)

    if date_str:
        hpf_xml = re.sub(
            r'(<opf:meta name="date" content="text")(?:/>|>(.*?)</opf:meta>)',
            rf'\1>{xml_escape(date_str)}</opf:meta>',
            hpf_xml)

    return hpf_xml

# ── Main ────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Convert Pandoc JSON AST to HWPX')
    parser.add_argument('--output', required=True, help='Output .hwpx file path')
    args = parser.parse_args()

    json_input = sys.stdin.read()
    ast = json.loads(json_input)

    meta = ast.get("meta", {})
    title = extract_meta_text(meta.get("title"))
    subtitle = extract_meta_text(meta.get("subtitle"))
    author = extract_meta_text(meta.get("author"))
    date_str = extract_meta_text(meta.get("date"))
    blocks = ast.get("blocks", [])

    script_dir = os.path.dirname(os.path.abspath(__file__))
    skeleton_path = os.path.join(script_dir, "Skeleton.hwpx")
    if not os.path.exists(skeleton_path):
        print(f"ERROR: Skeleton.hwpx not found at {skeleton_path}", file=sys.stderr)
        sys.exit(1)

    # Build section0.xml
    section_xml = build_section_xml(skeleton_path, blocks,
                                    title, subtitle, author, date_str)

    # Update header.xml and content.hpf
    with zipfile.ZipFile(skeleton_path, 'r') as z:
        hpf_xml = z.read('Contents/content.hpf').decode('utf-8')
        header_xml = z.read('Contents/header.xml').decode('utf-8')
    updated_hpf = update_content_hpf(hpf_xml, title, author, date_str)
    updated_header = update_header_xml(header_xml)

    # Repackage ZIP
    output_path = args.output
    with tempfile.NamedTemporaryFile(suffix='.hwpx', delete=False) as tmp:
        tmp_path = tmp.name

    try:
        with zipfile.ZipFile(skeleton_path, 'r') as src_zip:
            with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as dst_zip:
                for item in src_zip.infolist():
                    if item.filename == 'Contents/section0.xml':
                        dst_zip.writestr(item, section_xml.encode('utf-8'))
                    elif item.filename == 'Contents/content.hpf':
                        dst_zip.writestr(item, updated_hpf.encode('utf-8'))
                    elif item.filename == 'Contents/header.xml':
                        dst_zip.writestr(item, updated_header.encode('utf-8'))
                    else:
                        dst_zip.writestr(item, src_zip.read(item.filename))

        shutil.move(tmp_path, output_path)
        print(f"HWPX written to {output_path}", file=sys.stderr)

    except Exception:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise


if __name__ == '__main__':
    main()
