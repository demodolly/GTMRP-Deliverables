#!/usr/bin/env python3
"""Build the UIF single-slide summary — barriers & stakeholder impacts focus."""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from project_summary_content import PROJECTS

DECK = Path("/workspace/GTM Performance & Readiness Project Update - August 2026.pptx")
UIF_KEY = "Unified Intelligence Framework (UIF)"

NAVY = RGBColor(0x1F, 0x4E, 0x79)
DARK = RGBColor(0x33, 0x33, 0x33)
MID = RGBColor(0x55, 0x55, 0x55)
ACCENT = RGBColor(0xC0, 0x50, 0x00)
CONTEXT_FILL = RGBColor(0xE8, 0xEF, 0xF4)
RISK_FILL = RGBColor(0xFD, 0xF2, 0xE9)
STAKEHOLDER_FILL = RGBColor(0xFF, 0xF8, 0xE8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BORDER = RGBColor(0xCC, 0xCC, 0xCC)

# Barrier → stakeholders affected → impact (management-facing)
UIF_STAKEHOLDER_BARRIERS = [
    {
        "barrier": "One Cisco CRM freeze — no SFDC changes permitted FY27 H1",
        "stakeholders": "Sales, Orchestration, SFDC BCO",
        "impact": "Seller-facing Workfront strategic data cannot appear in CRM; Phase 2 value trapped in backend logs; sales cannot validate enriched lead quality or act on initiative/content context.",
    },
    {
        "barrier": "CCID & DTID still mandatory for SFDC lead creation, MSP & Last Touch reporting",
        "stakeholders": "Marketing, Reporting, GTMPR",
        "impact": "Teams must maintain dual manual governance (Workfront + legacy IDs); reporting cannot fully adopt UTM/Channel ID; attribution figures remain at risk of CCID mixing.",
    },
    {
        "barrier": "Incomplete Workfront adoption — not all teams onboarded",
        "stakeholders": "Marketing, CTT requestors, GTMPR",
        "impact": "Non-Workfront teams continue CTT legacy processes; UIF ROI capped; hybrid operating model persists longer than planned.",
    },
    {
        "barrier": "Adobe North Star & Reporting FY27 refresh schedule not locked",
        "stakeholders": "Reporting, DOPT, GTMPR",
        "impact": "Interim UIF field model may require rework; reporting teams delay model updates; North Star visibility deferred.",
    },
    {
        "barrier": "Tealium vs Adobe Capture undecided; Marketo/Eloqua path uncertain (TEA002/003)",
        "stakeholders": "MarTech BCOs, Tealium, Web Analytics",
        "impact": "URL/session capture logic in review; investment in one MAP stack may be wasted; re-work risk on in-flight tagging.",
    },
    {
        "barrier": "Business-Ready Workfront Dataset (GTMRP) not started — raw vs golden schema unresolved",
        "stakeholders": "Reporting, Data Engineering, Orchestration",
        "impact": "No single Workfront source of truth for attribution models; analysts continue ad-hoc joins; CTT bridge and UIF field definitions keep shifting.",
    },
    {
        "barrier": "Phase 1 incomplete — PUB002 AEM Content ID in progress; 39 reqs Not Started; UTM002 deferred",
        "stakeholders": "Publishing, Web, GTMPR, Phase 2 teams",
        "impact": "Content-level journey attribution blocked; all 12 Phase 2 requirements cannot start; programme timeline slips with every Phase 1 delay.",
    },
]


def _set_margin(tf, top=0.04, bottom=0.04, left=0.07, right=0.07):
    tf.margin_top = Inches(top)
    tf.margin_bottom = Inches(bottom)
    tf.margin_left = Inches(left)
    tf.margin_right = Inches(right)


def _add_run(paragraph, text, *, size=10, bold=False, color=DARK, italic=False):
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = "Arial"
    return run


def _box(slide, left, top, width, height, fill):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = BORDER
    shape.line.width = Pt(0.75)
    return shape


def _barrier_block(tf, item, *, size=8):
    p = tf.add_paragraph()
    p.space_before = Pt(6)
    p.space_after = Pt(1)
    _add_run(p, "▸ ", size=size, bold=True, color=ACCENT)
    _add_run(p, item["barrier"], size=size, bold=True, color=DARK)

    p2 = tf.add_paragraph()
    p2.space_after = Pt(1)
    _add_run(p2, "   Stakeholders: ", size=size - 1, bold=True, color=NAVY)
    _add_run(p2, item["stakeholders"], size=size - 1, color=MID)

    p3 = tf.add_paragraph()
    p3.space_after = Pt(4)
    _add_run(p3, "   Impact: ", size=size - 1, bold=True, color=ACCENT)
    _add_run(p3, item["impact"], size=size - 1, color=MID)


def build_uif_slide(prs: Presentation):
    """Single UIF slide: compact goal/benefits top; barriers & stakeholder impacts dominate."""
    layout_idx = 43 if len(prs.slide_layouts) > 43 else len(prs.slide_layouts) - 1
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])

    data = PROJECTS[UIF_KEY]
    sw = prs.slide_width
    sh = prs.slide_height
    m = Inches(0.22)

    # ── Header ────────────────────────────────────────────────────────────
    banner_h = Inches(0.95)
    banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, m, m, sw - 2 * m, banner_h)
    banner.fill.solid()
    banner.fill.fore_color.rgb = NAVY
    banner.line.fill.background()

    tf = banner.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    _set_margin(tf, 0.06, 0.05, 0.12, 0.12)

    p0 = tf.paragraphs[0]
    _add_run(p0, "Unified Intelligence Framework (UIF): Phase 1 + Phase 2", size=18, bold=True, color=WHITE)
    p1 = tf.add_paragraph()
    p1.space_before = Pt(3)
    _add_run(p1, "Goal: ", size=10, bold=True, color=WHITE)
    _add_run(
        p1,
        "Move from legacy CCID/DTID to automated Workfront Universal Key attribution across "
        "Marketo & Eloqua — without breaking FY27 reporting, lead creation, or sales operations.",
        size=10,
        color=WHITE,
    )

    # ── Context strip (combined programme + benefits) ───────────────────────
    ctx_top = m + banner_h + Inches(0.08)
    ctx_h = Inches(1.55)
    ctx_w = sw - 2 * m

    ctx_box = _box(slide, m, ctx_top, ctx_w, ctx_h, CONTEXT_FILL)
    tfc = ctx_box.text_frame
    tfc.word_wrap = True
    _set_margin(tfc, 0.05, 0.04, 0.1, 0.1)

    h = tfc.paragraphs[0]
    _add_run(h, "What We Are Delivering (Phase 1 + Phase 2 Combined)", size=10, bold=True, color=NAVY)

    p = tfc.add_paragraph()
    p.space_before = Pt(3)
    _add_run(
        p,
        "Phase 1 (now): UTM governance, dual ingestion of legacy + modern IDs, Snowflake-ready schema.  "
        "Phase 2 (next): Workfront auto-generates Channel & Content IDs, embedded on AEM pages & form submissions, "
        "flowing through MarTech to SFDC when permitted.",
        size=8,
        color=MID,
    )

    p = tfc.add_paragraph()
    p.space_before = Pt(5)
    _add_run(p, "Key Business Benefits: ", size=8, bold=True, color=NAVY)
    benefits_short = [
        "No big-bang cutover — reporting continuity during transition",
        "Accurate attribution at lead creation today",
        "Zero-latency Workfront-to-Marketo/Eloqua automation (Phase 2)",
        "Full content-journey insight & Adobe North Star alignment",
    ]
    _add_run(p, "  •  ".join(benefits_short), size=7.5, color=MID)

    # ── Main focus: Barriers & Stakeholder Impacts (~65% of slide) ────────
    risk_top = ctx_top + ctx_h + Inches(0.08)
    risk_h = sh - risk_top - m

    risk_box = _box(slide, m, risk_top, ctx_w, risk_h, RISK_FILL)
    tfr = risk_box.text_frame
    tfr.word_wrap = True
    _set_margin(tfr, 0.06, 0.05, 0.1, 0.1)

    h = tfr.paragraphs[0]
    _add_run(h, "⚠  ", size=13, bold=True, color=ACCENT)
    _add_run(h, "Barriers & Stakeholder Impacts — Why Delivery Is Constrained", size=12, bold=True, color=NAVY)

    p = tfr.add_paragraph()
    p.space_before = Pt(2)
    _add_run(
        p,
        "We are actively delivering UIF capabilities, but upstream dependencies and adoption gaps "
        "limit value realisation for key stakeholder groups:",
        size=8,
        italic=True,
        color=MID,
    )

    # Two columns of barrier blocks inside the risk panel
    col_gap = Inches(0.15)
    col_w = (ctx_w - col_gap) / 2
    left_items = UIF_STAKEHOLDER_BARRIERS[:4]
    right_items = UIF_STAKEHOLDER_BARRIERS[4:]

    # Left sub-column
    left_col = _box(
        slide,
        m + Inches(0.1),
        risk_top + Inches(0.62),
        col_w - Inches(0.05),
        risk_h - Inches(0.72),
        STAKEHOLDER_FILL,
    )
    tfl = left_col.text_frame
    tfl.word_wrap = True
    _set_margin(tfl, 0.04, 0.03, 0.06, 0.06)
    for item in left_items:
        _barrier_block(tfl, item, size=7.5)

    # Right sub-column
    right_col = _box(
        slide,
        m + Inches(0.1) + col_w + col_gap,
        risk_top + Inches(0.62),
        col_w - Inches(0.05),
        risk_h - Inches(0.72),
        STAKEHOLDER_FILL,
    )
    tfr2 = right_col.text_frame
    tfr2.word_wrap = True
    _set_margin(tfr2, 0.04, 0.03, 0.06, 0.06)
    for item in right_items:
        _barrier_block(tfr2, item, size=7.5)

    # Footer callout inside risk box
    p = tfr.add_paragraph()
    p.space_before = Pt(2)
    _add_run(p, "Management ask: ", size=8, bold=True, color=NAVY)
    _add_run(
        p,
        "Prioritise OCC engagement window, lock Reporting/North Star field catalog, ring-fence GTMRP "
        "foundation capacity, and complete Phase 1 publishing (PUB002) to unblock Phase 2 stakeholders.",
        size=7.5,
        color=MID,
    )

    return slide


def build_standalone(output: Path | None = None):
    base = Presentation(str(DECK))
    while len(base.slides) > 0:
        r_id = base.slides._sldIdLst[0]  # noqa: SLF001
        base.part.drop_rel(r_id.rId)
        del base.slides._sldIdLst[0]
    build_uif_slide(base)
    out = output or Path("/workspace/UIF_Summary_Slide.pptx")
    base.save(str(out))
    print(f"Standalone slide saved to {out}")


if __name__ == "__main__":
    build_standalone()
