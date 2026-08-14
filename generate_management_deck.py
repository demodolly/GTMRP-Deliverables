#!/usr/bin/env python3
"""Build management update deck from Project Summary + Requirements Register."""

from __future__ import annotations

from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Pt

from project_summary_content import (
    PROJECT_ORDER,
    PROJECTS,
    PROJECT_SHORT,
    REGISTER_SOURCE_KEYS,
)
from generate_uif_summary_slide import build_uif_slide

UIF_KEY = "Unified Intelligence Framework (UIF)"

REGISTER = Path("/workspace/BRD_Requirements_Register.xlsx")
SRC_DECK = Path("/workspace/GTM Performance & Readiness Project Update - August 2026.pptx")
OUT_DECK = Path("/workspace/GTM Performance & Readiness Project Update - August 2026.pptx")


def load_register_stats():
    wb = load_workbook(REGISTER, read_only=True)
    ws = wb["Requirements Register"]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    by_brd = defaultdict(Counter)
    for r in range(2, ws.max_row + 1):
        row = {h: ws.cell(r, c + 1).value for c, h in enumerate(headers)}
        if not row.get("Requirement ID"):
            continue
        proj = row["Project / BRD"]
        st = row.get("Status")
        by_brd[proj][str(st if st else "Not Started")] += 1
        by_brd[proj]["_total"] += 1

    # Roll up to portfolio projects
    portfolio = {}
    for key, sources in REGISTER_SOURCE_KEYS.items():
        combined = Counter()
        for src in sources:
            for k, v in by_brd[src].items():
                combined[k] += v
        portfolio[key] = combined
    return portfolio


def trim_slides_after(prs: Presentation, keep: int) -> None:
    while len(prs.slides) > keep:
        r_id = prs.slides._sldIdLst[keep]  # noqa: SLF001
        prs.part.drop_rel(r_id.rId)
        del prs.slides._sldIdLst[keep]


def set_title(slide, title: str, subtitle: str | None = None):
    if slide.shapes.title:
        slide.shapes.title.text = title
    if subtitle is not None:
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 1 and ph.has_text_frame:
                ph.text = subtitle
                break


def add_bullets(text_frame, items: list[str]):
    text_frame.clear()
    for i, item in enumerate(items):
        p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
        p.text = item
        p.level = 0
        p.font.size = Pt(15 if len(item) < 120 else 14)
        p.space_after = Pt(6)


def get_body_placeholder(slide):
    for shape in slide.placeholders:
        if shape.placeholder_format.idx == 1 and shape.has_text_frame:
            return shape.text_frame
    for shape in slide.shapes:
        if shape.has_text_frame and shape != slide.shapes.title:
            return shape.text_frame
    return None


def add_content_slide(prs, layout_idx, title, bullets):
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    set_title(slide, title)
    body = get_body_placeholder(slide)
    if body:
        add_bullets(body, bullets)
    return slide


def add_section_slide(prs, title, subtitle=""):
    slide = prs.slides.add_slide(prs.slide_layouts[12])
    set_title(slide, title, subtitle if subtitle else None)
    return slide


def status_line(stats: Counter) -> str:
    total = stats.get("_total", 0)
    ns = stats.get("Not Started", 0) + stats.get("blank", 0)
    tracked = total - ns
    complete = stats.get("Complete", 0)
    in_flight = stats.get("In Progress", 0) + stats.get("In Review", 0)
    return f"Register: {total} reqs | {complete} Complete | {in_flight} In Progress/In Review | {tracked} tracked"


def main():
    reg_stats = load_register_stats()

    prs = Presentation(str(SRC_DECK))
    keep = 3
    trim_slides_after(prs, keep)

    add_section_slide(prs, "Programme Overview", "GTM Performance & Readiness — August 2026")

    total_reqs = sum(s.get("_total", 0) for s in reg_stats.values())
    total_complete = sum(s.get("Complete", 0) for s in reg_stats.values())
    total_in_prog = sum(
        s.get("In Progress", 0) + s.get("In Review", 0) for s in reg_stats.values()
    )

    add_content_slide(
        prs,
        19,
        "Portfolio context — 5 programmes, delivering under upstream uncertainty",
        [
            f"{total_reqs} requirements across 5 portfolio programmes ({total_complete} Complete, {total_in_prog} In Progress/In Review)",
            "Programmes: UIF (Phase 1+2) | DTID & CTT Alignment | Workfront–CTT POC | Business-Ready Dataset | CTT Decommission",
            "Roadblocks from One Cisco CRM, Adobe North Star, and Business-Ready Dataset uncertainty force repeated pivots",
            "Each section: Description → Goal → Benefits → Barriers & Business Impacts",
        ],
    )

    for key in PROJECT_ORDER:
        data = PROJECTS[key]
        short = PROJECT_SHORT.get(key, key)
        subtitle = status_line(reg_stats.get(key, Counter()))

        if key == UIF_KEY:
            # Single three-zone summary slide (Phase 1 | Phase 2 | Risks)
            build_uif_slide(prs)
            continue

        add_section_slide(prs, short, subtitle)

        add_content_slide(prs, 19, f"{short} — Project Description", [data["description"]])
        add_content_slide(prs, 19, f"{short} — Project Goal", [data["goal"]])
        add_content_slide(prs, 19, f"{short} — Benefits to the Business", data["benefits"])

        barrier_bullets = []
        for pair in data["barrier_impacts"]:
            barrier_bullets.append(f"Barrier: {pair['barrier']}")
            barrier_bullets.append(f"→ Impact: {pair['impact']}")

        add_content_slide(prs, 19, f"{short} — Barriers & Business Impacts", barrier_bullets)

    add_section_slide(prs, "Appendix", "Living documents")
    add_content_slide(
        prs,
        19,
        "Refresh as the register grows",
        [
            "Project Summary.docx — run: python3 generate_project_summary.py",
            "Management deck — run: python3 generate_management_deck.py",
            "Requirements Register — underlying BRD-level detail (208 reqs across 7 BRDs)",
            "Portfolio view combines related BRDs — update project_summary_content.py to adjust narrative",
        ],
    )

    prs.save(str(OUT_DECK))
    print(f"Saved deck with {len(prs.slides)} slides to {OUT_DECK}")


if __name__ == "__main__":
    main()
