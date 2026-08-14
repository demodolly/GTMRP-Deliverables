#!/usr/bin/env python3
"""Build single-slide project summaries (UIF, CTT Alignment, etc.)."""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation

from summary_slide_builder import build_summary_slide
from summary_slide_content import (
    CTT_ALIGNMENT_SLIDE,
    UIF_SLIDE,
    UIF_TO_CTT_PIVOT_SLIDE,
)

DECK = Path("/workspace/GTM Performance & Readiness Project Update - August 2026.pptx")


def build_uif_slide(prs: Presentation):
    return build_summary_slide(prs, UIF_SLIDE)


def build_ctt_alignment_slide(prs: Presentation):
    return build_summary_slide(prs, CTT_ALIGNMENT_SLIDE)


def build_pivot_slide(prs: Presentation):
    return build_summary_slide(prs, UIF_TO_CTT_PIVOT_SLIDE)


def build_standalone(content, output: Path):
    base = Presentation(str(DECK))
    while len(base.slides) > 0:
        r_id = base.slides._sldIdLst[0]  # noqa: SLF001
        base.part.drop_rel(r_id.rId)
        del base.slides._sldIdLst[0]
    build_summary_slide(base, content)
    base.save(str(output))
    print(f"Standalone slide saved to {output}")


if __name__ == "__main__":
    build_standalone(UIF_SLIDE, Path("/workspace/UIF_Summary_Slide.pptx"))
    build_standalone(CTT_ALIGNMENT_SLIDE, Path("/workspace/CTT_Attribution_Data_Alignment_Summary_Slide.pptx"))
    build_standalone(UIF_TO_CTT_PIVOT_SLIDE, Path("/workspace/UIF_to_CTT_Pivot_Summary_Slide.pptx"))
