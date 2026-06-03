"""
Grade Report Generator
Reads Grades.xlsx and produces a professional PDF report per student.
"""

import os
import re
from collections import defaultdict
from datetime import date

import openpyxl
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    HRFlowable,
    KeepTogether,
)
from reportlab.platypus.flowables import Flowable

# ── Colour palette – Simcoe County District School Board brand colours ────────
# SCDSB dark navy  – header bar, borders
SCHOOL_BLUE = colors.HexColor("#003F5A")
SCHOOL_MID = colors.HexColor("#00476C")   # SCDSB medium blue – section headers
# SCDSB light blue  – alternating rows / accents
LIGHT_BG = colors.HexColor("#B1D2E1")
# tint of light blue – subtle row fill
LIGHT_BG_PALE = colors.HexColor("#DFF0F6")
WHITE = colors.white
BLACK = colors.HexColor("#1A1A1A")

# ── Grade thresholds ──────────────────────────────────────────────────────────


def letter_grade(pct: float) -> str:
    if pct >= 0.95:
        return "A+"
    if pct >= 0.87:
        return "A"
    if pct >= 0.80:
        return "A−"
    if pct >= 0.77:
        return "B+"
    if pct >= 0.73:
        return "B"
    if pct >= 0.70:
        return "B−"
    if pct >= 0.67:
        return "C+"
    if pct >= 0.63:
        return "C"
    if pct >= 0.60:
        return "C−"
    if pct >= 0.57:
        return "D+"
    if pct >= 0.53:
        return "D"
    if pct >= 0.50:
        return "D−"
    return "F"


def grade_color(pct: float) -> colors.Color:
    if pct >= 0.70:
        return colors.HexColor("#1A7340")   # green  – B− to A+
    if pct >= 0.60:
        return colors.HexColor("#B8860B")   # yellow – C range
    if pct >= 0.50:
        return colors.HexColor("#C85A00")   # orange – D range
    return colors.HexColor("#B22222")       # red    – F


# ── Module-level paragraph styles (#10) ──────────────────────────────────────
# Defined once here rather than re-created on every build_student_report call.

_STYLE_NAME = ParagraphStyle(
    "StudentName",
    fontSize=20, fontName="Helvetica-Bold",
    textColor=SCHOOL_BLUE, spaceAfter=2, alignment=TA_LEFT,
)
_STYLE_META = ParagraphStyle(
    "Meta",
    fontSize=9, fontName="Helvetica",
    textColor=SCHOOL_MID, spaceAfter=2, alignment=TA_LEFT,
)
_STYLE_META_R = ParagraphStyle(
    "MetaR", parent=_STYLE_META, alignment=TA_RIGHT,
)
_STYLE_SECTION = ParagraphStyle(
    "SectionHead",
    fontSize=11, fontName="Helvetica-Bold",
    textColor=WHITE, spaceAfter=0, spaceBefore=6,
    alignment=TA_LEFT, leftIndent=4,
)
_STYLE_NOTE = ParagraphStyle(
    "Note",
    fontSize=8, fontName="Helvetica-Oblique",
    textColor=SCHOOL_MID, alignment=TA_CENTER, spaceBefore=4,
)
_STYLE_SUM_LABEL = ParagraphStyle(
    "SumLabel",
    fontSize=9, fontName="Helvetica-Bold",
    textColor=SCHOOL_BLUE, alignment=TA_LEFT,
)
_STYLE_SUM_LABEL_LEAD = ParagraphStyle(
    "SumLbl", parent=_STYLE_SUM_LABEL, leading=13,
)
_STYLE_COL_HDR = ParagraphStyle(
    "ColHdr", fontSize=8, fontName="Helvetica-Bold",
    textColor=SCHOOL_BLUE, alignment=TA_LEFT,
)
_STYLE_COL_HDR_C = ParagraphStyle(
    "ColHdrC", parent=_STYLE_COL_HDR, alignment=TA_CENTER,
)
_STYLE_COL_HDR_C2 = ParagraphStyle(
    "ColHdrC2", parent=_STYLE_COL_HDR, alignment=TA_CENTER,
)
_STYLE_COL_HDR_R = ParagraphStyle(
    "ColHdrR", parent=_STYLE_COL_HDR, alignment=TA_CENTER,
)
# Reusable row styles (shared across all rows — no per-row unique names)
_STYLE_ROW_NAME = ParagraphStyle(
    "RowName", fontSize=8.5, fontName="Helvetica",
    textColor=BLACK, alignment=TA_LEFT,
)
_STYLE_ROW_VAL = ParagraphStyle(
    "RowVal", fontSize=8.5, fontName="Helvetica-Bold", alignment=TA_CENTER,
)
_STYLE_MOD_AVG_ROW = ParagraphStyle(
    "ModAvgRow", fontSize=8.5, fontName="Helvetica-Bold",
    textColor=SCHOOL_MID, alignment=TA_RIGHT,
)
_STYLE_MOD_AVG_VAL = ParagraphStyle(
    "ModAvgVal", fontSize=8.5, fontName="Helvetica-Bold",
    textColor=SCHOOL_MID, alignment=TA_CENTER,
)

# ── Summary card / scale / module-header styles (hoisted to avoid per-call alloc) ─
_STYLE_OV_PCT = ParagraphStyle(
    "OvPct", fontSize=22, fontName="Helvetica-Bold",
    alignment=TA_CENTER, leading=22, spaceAfter=0, spaceBefore=0,
)
_STYLE_OV_GRADE = ParagraphStyle(
    "OvGrade", fontSize=22, fontName="Helvetica-Bold",
    alignment=TA_CENTER, leading=22, spaceAfter=0, spaceBefore=0,
)
_STYLE_OV_CNT = ParagraphStyle(
    "OvCnt", fontSize=9, fontName="Helvetica",
    textColor=SCHOOL_MID, leading=12, alignment=TA_CENTER,
)
_STYLE_OV_CNT_VAL = ParagraphStyle(
    "OvCntVal", fontSize=16, fontName="Helvetica-Bold",
    textColor=SCHOOL_MID, alignment=TA_CENTER,
)
_STYLE_SCALE = ParagraphStyle(
    "Scale", fontSize=7.5, fontName="Helvetica", alignment=TA_CENTER,
)
_STYLE_MOD_HDR_AVG = ParagraphStyle(
    "ModHdrAvg", fontSize=10, fontName="Helvetica-Bold",
    textColor=WHITE, alignment=TA_RIGHT, rightIndent=4,
)


SCHOOL_NAME = "Westfield Academy"
REPORT_TITLE = "Student Grade Report"
REPORT_PERIOD = "Academic Year 2025–2026"


def _make_header_footer(school_name: str = SCHOOL_NAME):
    """Return a page-callback that draws the header/footer with the given school name."""
    def _callback(canvas, doc):
        canvas.saveState()
        w, h = LETTER

        # ── Top banner ────────────────────────────────────────────────────────
        canvas.setFillColor(SCHOOL_BLUE)
        canvas.rect(0, h - 0.75 * inch, w, 0.75 * inch, fill=1, stroke=0)

        canvas.setFillColor(LIGHT_BG)
        canvas.rect(0, h - 0.78 * inch, w, 0.03 * inch, fill=1, stroke=0)

        canvas.setFont("Helvetica-Bold", 15)
        canvas.setFillColor(WHITE)
        canvas.drawCentredString(w / 2, h - 0.47 * inch, school_name)
        canvas.setFont("Helvetica", 9)
        canvas.drawCentredString(w / 2, h - 0.64 * inch,
                                 f"{REPORT_TITLE}  ·  {REPORT_PERIOD}")

        # ── Bottom footer ─────────────────────────────────────────────────────
        canvas.setFillColor(SCHOOL_BLUE)
        canvas.rect(0, 0, w, 0.45 * inch, fill=1, stroke=0)
        canvas.setFillColor(LIGHT_BG)
        canvas.rect(0, 0.42 * inch, w, 0.03 * inch, fill=1, stroke=0)

        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(WHITE)
        canvas.drawString(0.5 * inch, 0.17 * inch,
                          f"Generated: {date.today().strftime('%B %d, %Y')}")
        canvas.drawCentredString(w / 2, 0.17 * inch, school_name)
        canvas.drawRightString(w - 0.5 * inch, 0.17 * inch,
                               f"Page {doc.page}")

        canvas.restoreState()
    return _callback


# ── Helper: percentage bar as a mini-table cell ───────────────────────────────
def pct_bar_table(pct: float, bar_width: float = 1.1 * inch) -> Table:
    """Return a tiny 1-row table that looks like a progress bar."""
    filled = max(0.0, min(1.0, pct)) * bar_width
    empty = bar_width - filled
    bar_color = grade_color(pct)
    t = Table(
        [["", ""]],
        colWidths=[filled, empty],
        rowHeights=[7],
    )
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), bar_color),
        ("BACKGROUND", (1, 0), (1, 0), LIGHT_BG_PALE),
        ("LINEABOVE",  (0, 0), (-1, 0), 0.5, LIGHT_BG),
        ("LINEBELOW",  (0, 0), (-1, 0), 0.5, LIGHT_BG),
        ("LEFTPADDING",  (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING",   (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return t


# ── Main builder ──────────────────────────────────────────────────────────────
def build_student_report(student: dict, output_path: str, school_name: str = SCHOOL_NAME,
                         teacher_name: str = "", teacher_email: str = ""):
    """
    student = {
        'id': ...,
        'last': ...,
        'first': ...,
        'modules': {
            '1': [('1.1 Discussion', 0.67), ...],
            '2': [('2.1 Discussion', 0.94), ...],
        }
    }
    """
    # Use module-level style constants (#10)
    style_name = _STYLE_NAME
    style_meta = _STYLE_META
    style_section = _STYLE_SECTION
    style_note = _STYLE_NOTE
    style_summary_label = _STYLE_SUM_LABEL

    doc = BaseDocTemplate(
        output_path,
        pagesize=LETTER,
        rightMargin=0.6 * inch,
        leftMargin=0.6 * inch,
        topMargin=1.0 * inch,
        bottomMargin=0.65 * inch,
    )

    content_width = doc.width   # usable width between margins

    frame = Frame(
        doc.leftMargin, doc.bottomMargin,
        doc.width, doc.height,
        id="main"
    )
    template = PageTemplate(id="school", frames=[frame],
                            onPage=_make_header_footer(school_name))
    doc.addPageTemplates([template])

    story = []

    # ── Student identity card ────────────────────────────────────────────────
    full_name = f"{student['first']} {student['last']}"
    story.append(Paragraph(full_name, style_name))
    story.append(Spacer(1, 6))
    meta_table = Table(
        [[
            Paragraph(f"Report Date: <b>{date.today().strftime('%B %d, %Y')}</b>",
                      _STYLE_META_R),
        ]],
        colWidths=[content_width],
        hAlign="LEFT",
    )
    meta_table.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=2,
                            color=LIGHT_BG, spaceAfter=8))

    # ── Overall summary ──────────────────────────────────────────────────────
    _all_sections = list(student["modules"].values()) + \
        list(student.get("journals", {}).values())
    submitted_count = sum(
        1 for mod in _all_sections
        for _, s in mod if (isinstance(s, float) and s > 0.0) or s == "?"
    )
    total_assignments = sum(len(a) for a in _all_sections)

    # Use weighted grade from Column D; fall back to computed average if absent
    all_scores = [score for mod in _all_sections
                  for _, score in mod if isinstance(score, float) and score > 0.0]
    fallback_avg = sum(all_scores) / len(all_scores) if all_scores else 0.0
    display_grade = student.get("weighted_grade")
    if display_grade is None:
        display_grade = fallback_avg
    display_pct = f"{display_grade * 100:.1f}%"
    display_letter = letter_grade(display_grade)
    ov_color = grade_color(display_grade)

    summary_data = [
        [
            Paragraph(
                'Current Grade',
                _STYLE_SUM_LABEL_LEAD,
            ),
            Paragraph(
                f'<font color="{ov_color.hexval()}">{display_pct}</font>',
                _STYLE_OV_PCT),
            Paragraph(
                f'<font color="{ov_color.hexval()}">{display_letter}</font>',
                _STYLE_OV_GRADE),
            Paragraph("Assignments<br/>Submitted", _STYLE_OV_CNT),
            Paragraph(f"{submitted_count}/{total_assignments}",
                      _STYLE_OV_CNT_VAL),
        ]
    ]
    summary_table = Table(
        summary_data,
        colWidths=[
            content_width * 0.28,
            content_width * 0.18,
            content_width * 0.18,
            content_width * 0.22,
            content_width * 0.14,
        ],
        rowHeights=[44],
    )
    summary_table.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, -1), LIGHT_BG_PALE),
        ("ROUNDEDCORNERS", [4]),
        ("BOX",          (0, 0), (-1, -1), 1.5, SCHOOL_BLUE),
        ("LINEBEFORE",   (1, 0), (1, 0), 1, LIGHT_BG),
        ("LINEBEFORE",   (2, 0), (2, 0), 1, LIGHT_BG),
        ("LINEBEFORE",   (3, 0), (3, 0), 1, LIGHT_BG),
        ("LINEBEFORE",   (4, 0), (4, 0), 1, LIGHT_BG),
        ("ALIGN",        (1, 0), (2, 0), "CENTER"),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING",   (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 10))

    # ── Grade scale legend ───────────────────────────────────────────────────
    scale_items = [
        ("A+/A/A−",  "80–100%", colors.HexColor("#1A7340")),
        ("B+/B/B−",  "70–79%",  colors.HexColor("#1A7340")),
        ("C+/C/C−",  "60–69%",  colors.HexColor("#B8860B")),
        ("D+/D/D−",  "50–59%",  colors.HexColor("#C85A00")),
        ("F",         "< 50%",   colors.HexColor("#B22222")),
    ]
    scale_cells = []
    for lg, rng, col in scale_items:
        scale_cells.append(
            Paragraph(f'<font color="{col.hexval()}"><b>{lg}</b></font>  {rng}',
                      _STYLE_SCALE)
        )
    scale_table = Table([scale_cells],
                        colWidths=[content_width / 5] * 5,
                        rowHeights=[14])
    scale_table.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, -1), LIGHT_BG_PALE),
        ("BOX",          (0, 0), (-1, -1), 0.5, LIGHT_BG),
        ("INNERGRID",    (0, 0), (-1, -1), 0.5, LIGHT_BG),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",   (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(scale_table)
    story.append(Spacer(1, 12))

    # ── Per-module sections ───────────────────────────────────────────────────
    module_nums = sorted(student["modules"].keys(),
                         key=lambda x: float(x) if x.replace(".", "").isdigit() else x)

    for mod_num in module_nums:
        assignments = student["modules"][mod_num]

        # Module average (0% scores treated as not submitted; ? treated as pending)
        valid_scores = [
            s for _, s in assignments if isinstance(s, float) and s > 0.0]
        mod_avg = sum(valid_scores) / \
            len(valid_scores) if valid_scores else 0.0

        # ── Section header bar ────────────────────────────────────────────────
        header_data = [[
            Paragraph(f"  Module {mod_num}", style_section),
            Paragraph(
                f'Avg: <b>{mod_avg * 100:.1f}%</b>  '
                f'({letter_grade(mod_avg)})',
                _STYLE_MOD_HDR_AVG
            ),
        ]]
        header_table = Table(
            header_data,
            colWidths=[content_width * 0.6, content_width * 0.4],
            rowHeights=[20],
        )
        header_table.setStyle(TableStyle([
            ("BACKGROUND",   (0, 0), (-1, -1), SCHOOL_MID),
            ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING",  (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING",   (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))

        # ── Assignment rows ───────────────────────────────────────────────────
        col_widths = [
            content_width * 0.40,   # Assignment name
            content_width * 0.15,   # Score %
            content_width * 0.10,   # Grade
            content_width * 0.35,   # Bar
        ]

        # Column headers
        table_rows = [[
            Paragraph("Assignment", _STYLE_COL_HDR),
            Paragraph("Score",      _STYLE_COL_HDR_C),
            Paragraph("Grade",      _STYLE_COL_HDR_C2),
            Paragraph("Visual",     _STYLE_COL_HDR_R),
        ]]

        row_styles = [
            # Header row styling
            ("BACKGROUND",   (0, 0), (-1, 0), LIGHT_BG),
            ("LINEBELOW",    (0, 0), (-1, 0), 1, SCHOOL_BLUE),
            ("FONTNAME",     (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE",     (0, 0), (-1, 0), 8),
            ("TOPPADDING",   (0, 0), (-1, 0), 3),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 3),
        ]

        for i, (asgn_name, score) in enumerate(assignments):
            row_idx = i + 1
            bg = LIGHT_BG_PALE if row_idx % 2 == 0 else WHITE
            is_not_submitted = False
            is_submitted_pending = False

            if isinstance(score, float) and score > 0.0:
                pct_str = f"{score * 100:.1f}%"
                lg = letter_grade(score)
                gc = grade_color(score)
                val_para = Paragraph(
                    f'<font color="{gc.hexval()}">{pct_str}</font>', _STYLE_ROW_VAL
                )
                grade_para = Paragraph(
                    f'<font color="{gc.hexval()}"><b>{lg}</b></font>', _STYLE_ROW_VAL
                )
                bar = pct_bar_table(score, bar_width=col_widths[3] - 16)
            elif score == "?":
                val_para = Paragraph(
                    '<font size="7" color="#B8860B">Submitted</font>', _STYLE_ROW_VAL)
                grade_para = Paragraph("", _STYLE_ROW_VAL)
                bar = Paragraph("",    _STYLE_ROW_VAL)
                is_submitted_pending = True
            else:
                val_para = Paragraph(
                    '<font size="7">Not Submitted</font>', _STYLE_ROW_VAL)
                grade_para = Paragraph("", _STYLE_ROW_VAL)
                bar = Paragraph("",    _STYLE_ROW_VAL)
                is_not_submitted = True

            table_rows.append([
                Paragraph(asgn_name, _STYLE_ROW_NAME),
                val_para,
                grade_para,
                bar,
            ])

            row_styles += [
                ("BACKGROUND",   (0, row_idx), (-1, row_idx), bg),
                ("TOPPADDING",   (0, row_idx), (-1, row_idx), 4),
                ("BOTTOMPADDING", (0, row_idx), (-1, row_idx), 4),
                ("LEFTPADDING",  (0, row_idx), (-1, row_idx), 5),
                ("RIGHTPADDING", (0, row_idx), (-1, row_idx), 5),
                ("VALIGN",       (0, row_idx), (-1, row_idx), "MIDDLE"),
            ]
            if is_not_submitted or is_submitted_pending:
                row_styles.append(("SPAN", (1, row_idx), (2, row_idx)))

        # Module average footer row
        table_rows.append([
            Paragraph("Module Average", _STYLE_MOD_AVG_ROW),
            Paragraph(f"{mod_avg * 100:.1f}%", _STYLE_MOD_AVG_VAL),
            Paragraph(letter_grade(mod_avg),    _STYLE_MOD_AVG_VAL),
            pct_bar_table(mod_avg, bar_width=col_widths[3] - 16),
        ])
        footer_idx = len(table_rows) - 1
        row_styles += [
            ("BACKGROUND",   (0, footer_idx), (-1, footer_idx),
             LIGHT_BG),
            ("LINEABOVE",    (0, footer_idx), (-1, footer_idx), 1, SCHOOL_BLUE),
            ("FONTNAME",     (0, footer_idx), (-1, footer_idx), "Helvetica-Bold"),
            ("TOPPADDING",   (0, footer_idx), (-1, footer_idx), 4),
            ("BOTTOMPADDING", (0, footer_idx), (-1, footer_idx), 4),
            ("LEFTPADDING",  (0, footer_idx), (-1, footer_idx), 5),
            ("RIGHTPADDING", (0, footer_idx), (-1, footer_idx), 5),
            ("VALIGN",       (0, footer_idx), (-1, footer_idx), "MIDDLE"),
        ]

        row_styles += [
            ("BOX",       (0, 0), (-1, -1), 1, SCHOOL_BLUE),
            ("LINEBELOW", (0, -1), (-1, -1), 1, SCHOOL_BLUE),
        ]

        assign_table = Table(table_rows, colWidths=col_widths)
        assign_table.setStyle(TableStyle(row_styles))

        story.append(KeepTogether([header_table, assign_table]))
        story.append(Spacer(1, 10))

    # ── Journals section ─────────────────────────────────────────────────────
    journals = student.get("journals", {})
    if journals:
        journal_nums = sorted(
            journals.keys(), key=lambda x: int(x) if x.isdigit() else x
        )

        # Top-level "Journals" banner (SCHOOL_BLUE to distinguish from module headers)
        journals_banner_data = [[Paragraph("  Journals", style_section)]]
        journals_banner = Table(
            journals_banner_data,
            colWidths=[content_width],
            rowHeights=[20],
        )
        journals_banner.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), SCHOOL_BLUE),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING",   (0, 0), (-1, -1), 0),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
            ("TOPPADDING",    (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))

        for idx, j_num in enumerate(journal_nums):
            j_assignments = journals[j_num]
            j_valid = [s for _, s in j_assignments if isinstance(
                s, float) and s > 0.0]
            j_avg = sum(j_valid) / len(j_valid) if j_valid else 0.0

            j_header_data = [[
                Paragraph(f"  Journal {j_num}", style_section),
                Paragraph(
                    f'Avg: <b>{j_avg * 100:.1f}%</b>  ({letter_grade(j_avg)})',
                    _STYLE_MOD_HDR_AVG,
                ),
            ]]
            j_header_table = Table(
                j_header_data,
                colWidths=[content_width * 0.6, content_width * 0.4],
                rowHeights=[20],
            )
            j_header_table.setStyle(TableStyle([
                ("BACKGROUND",    (0, 0), (-1, -1), SCHOOL_MID),
                ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING",   (0, 0), (-1, -1), 0),
                ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
                ("TOPPADDING",    (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]))

            j_col_widths = [
                content_width * 0.40,
                content_width * 0.15,
                content_width * 0.10,
                content_width * 0.35,
            ]
            j_table_rows = [[
                Paragraph("Assignment", _STYLE_COL_HDR),
                Paragraph("Score",      _STYLE_COL_HDR_C),
                Paragraph("Grade",      _STYLE_COL_HDR_C2),
                Paragraph("Visual",     _STYLE_COL_HDR_R),
            ]]
            j_row_styles = [
                ("BACKGROUND",    (0, 0), (-1, 0), LIGHT_BG),
                ("LINEBELOW",     (0, 0), (-1, 0), 1, SCHOOL_BLUE),
                ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE",      (0, 0), (-1, 0), 8),
                ("TOPPADDING",    (0, 0), (-1, 0), 3),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 3),
            ]

            for i, (asgn_name, score) in enumerate(j_assignments):
                row_idx = i + 1
                bg = LIGHT_BG_PALE if row_idx % 2 == 0 else WHITE
                is_not_submitted = False
                is_submitted_pending = False

                if isinstance(score, float) and score > 0.0:
                    pct_str = f"{score * 100:.1f}%"
                    lg = letter_grade(score)
                    gc = grade_color(score)
                    val_para = Paragraph(
                        f'<font color="{gc.hexval()}">{pct_str}</font>', _STYLE_ROW_VAL
                    )
                    grade_para = Paragraph(
                        f'<font color="{gc.hexval()}"><b>{lg}</b></font>', _STYLE_ROW_VAL
                    )
                    bar = pct_bar_table(score, bar_width=j_col_widths[3] - 16)
                elif score == "?":
                    val_para = Paragraph(
                        '<font size="7" color="#B8860B">Submitted</font>', _STYLE_ROW_VAL)
                    grade_para = Paragraph("", _STYLE_ROW_VAL)
                    bar = Paragraph("", _STYLE_ROW_VAL)
                    is_submitted_pending = True
                else:
                    val_para = Paragraph(
                        '<font size="7">Not Submitted</font>', _STYLE_ROW_VAL)
                    grade_para = Paragraph("", _STYLE_ROW_VAL)
                    bar = Paragraph("", _STYLE_ROW_VAL)
                    is_not_submitted = True

                j_table_rows.append([
                    Paragraph(asgn_name, _STYLE_ROW_NAME),
                    val_para,
                    grade_para,
                    bar,
                ])
                j_row_styles += [
                    ("BACKGROUND",    (0, row_idx), (-1, row_idx), bg),
                    ("TOPPADDING",    (0, row_idx), (-1, row_idx), 4),
                    ("BOTTOMPADDING", (0, row_idx), (-1, row_idx), 4),
                    ("LEFTPADDING",   (0, row_idx), (-1, row_idx), 5),
                    ("RIGHTPADDING",  (0, row_idx), (-1, row_idx), 5),
                    ("VALIGN",        (0, row_idx), (-1, row_idx), "MIDDLE"),
                ]
                if is_not_submitted or is_submitted_pending:
                    j_row_styles.append(("SPAN", (1, row_idx), (2, row_idx)))

            j_table_rows.append([
                Paragraph("Journal Average", _STYLE_MOD_AVG_ROW),
                Paragraph(f"{j_avg * 100:.1f}%", _STYLE_MOD_AVG_VAL),
                Paragraph(letter_grade(j_avg),    _STYLE_MOD_AVG_VAL),
                pct_bar_table(j_avg, bar_width=j_col_widths[3] - 16),
            ])
            j_footer_idx = len(j_table_rows) - 1
            j_row_styles += [
                ("BACKGROUND",    (0, j_footer_idx), (-1, j_footer_idx), LIGHT_BG),
                ("LINEABOVE",     (0, j_footer_idx),
                 (-1, j_footer_idx), 1, SCHOOL_BLUE),
                ("FONTNAME",      (0, j_footer_idx),
                 (-1, j_footer_idx), "Helvetica-Bold"),
                ("TOPPADDING",    (0, j_footer_idx), (-1, j_footer_idx), 4),
                ("BOTTOMPADDING", (0, j_footer_idx), (-1, j_footer_idx), 4),
                ("LEFTPADDING",   (0, j_footer_idx), (-1, j_footer_idx), 5),
                ("RIGHTPADDING",  (0, j_footer_idx), (-1, j_footer_idx), 5),
                ("VALIGN",        (0, j_footer_idx), (-1, j_footer_idx), "MIDDLE"),
            ]
            j_row_styles += [
                ("BOX",       (0, 0), (-1, -1), 1, SCHOOL_BLUE),
                ("LINEBELOW", (0, -1), (-1, -1), 1, SCHOOL_BLUE),
            ]

            j_assign_table = Table(j_table_rows, colWidths=j_col_widths)
            j_assign_table.setStyle(TableStyle(j_row_styles))

            # Keep the banner with the first journal group so it never floats alone
            if idx == 0:
                story.append(KeepTogether(
                    [journals_banner, j_header_table, j_assign_table]))
            else:
                story.append(KeepTogether([j_header_table, j_assign_table]))
            story.append(Spacer(1, 10))

    # ── Closing note ─────────────────────────────────────────────────────────
    story.append(HRFlowable(width="100%", thickness=1,
                            color=LIGHT_BG, spaceBefore=4, spaceAfter=4))
    if teacher_name and teacher_email:
        contact = (
            f"Please contact {full_name}'s teacher {teacher_name} "
            f"at {teacher_email} if you have any questions or concerns."
        )
    elif teacher_name:
        contact = (
            f"Please contact {full_name}'s teacher {teacher_name} "
            f"if you have any questions or concerns."
        )
    else:
        contact = f"Please contact {full_name}'s teacher if you have any questions."
    story.append(Paragraph(
        contact,
        style_note,
    ))

    doc.build(story)


# ── Data loading ──────────────────────────────────────────────────────────────
def get_module_number(assignment_name: str) -> str:
    """Extract the leading numeric module prefix, e.g. '1.2 Quiz' → '1'."""
    m = re.match(r"^(\d+)\.", assignment_name.strip())
    if m:
        return m.group(1)
    return "Other"


def get_journal_number(assignment_name: str) -> str | None:
    """If assignment uses journal notation (J1.1, J2.3, etc.) return the group number."""
    m = re.match(r"^[Jj](\d+)\.", assignment_name.strip())
    if m:
        return m.group(1)
    return None


def load_grades(xlsx_path: str) -> list[dict]:
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    header = rows[0]
    # Columns: 0=StudentID, 1=Last, 2=First, 3=empty, 4+=assignments
    assignment_cols = [
        (col_idx, str(col_name).strip())
        for col_idx, col_name in enumerate(header)
        if col_idx >= 4 and col_name is not None
    ]

    students = []
    for row in rows[1:]:
        if row[0] is None:
            continue

        student_id = row[0]
        last_name = str(row[1]).strip() if row[1] else ""
        first_name = str(row[2]).strip() if row[2] else ""

        # Column D (index 3) — weighted grade
        weighted_grade: float | None = None
        weighted_raw = row[3] if len(row) > 3 else None
        if isinstance(weighted_raw, (int, float)):
            wv = float(weighted_raw)
            weighted_grade = wv / 100.0 if wv > 1.0 else wv
        elif weighted_raw is not None:
            raw_w = str(weighted_raw).strip().rstrip("%")
            try:
                wv = float(raw_w)
                weighted_grade = wv / 100.0 if wv > 1.0 else wv
            except (ValueError, TypeError):
                weighted_grade = None

        modules: dict[str, list] = defaultdict(list)
        journals: dict[str, list] = defaultdict(list)
        for col_idx, asgn_name in assignment_cols:
            score = row[col_idx] if col_idx < len(row) else None
            if isinstance(score, (int, float)):
                score = float(score)
            elif score is not None:
                raw = str(score).strip()
                if raw.rstrip("% ") == "?":
                    score = "?"
                else:
                    try:
                        score = float(score)
                    except (ValueError, TypeError):
                        score = None
            jnum = get_journal_number(asgn_name)
            if jnum is not None:
                journals[jnum].append((asgn_name, score))
            else:
                mod_num = get_module_number(asgn_name)
                modules[mod_num].append((asgn_name, score))

        students.append({
            "id":            student_id,
            "last":          last_name,
            "first":         first_name,
            "weighted_grade": weighted_grade,
            "modules":       dict(modules),
            "journals":      dict(journals),
        })

    return students


# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    xlsx_path = os.path.join(script_dir, "Grades.xlsx")
    output_dir = os.path.join(script_dir, "reports")
    os.makedirs(output_dir, exist_ok=True)

    print(f"Reading: {xlsx_path}")
    students = load_grades(xlsx_path)

    if not students:
        print("No student data found.")
        return

    for student in students:
        filename = (
            f"{student['last']}_{student['first']}.pdf"
            .replace(" ", "_")
        )
        out_path = os.path.join(output_dir, filename)
        print(f"  Generating → {filename}")
        build_student_report(student, out_path)

    print(f"\nDone. {len(students)} report(s) saved to: {output_dir}")


if __name__ == "__main__":
    main()
