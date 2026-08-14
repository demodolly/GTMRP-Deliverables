"""Speaker notes for GTM Performance & Readiness Project Update deck (August 2026)."""

from __future__ import annotations

from summary_slide_content import (
    CTT_ALIGNMENT_SLIDE,
    UIF_SLIDE,
    UIF_TO_CTT_PIVOT_SLIDE,
)
from summary_slide_builder import SummarySlideContent

# Notes keyed by 1-based slide index (25 slides in current deck)


def _notes_from_summary(content: SummarySlideContent, *, extra_open: str = "", extra_close: str = "") -> str:
    lines = []
    if extra_open:
        lines.append(extra_open)
        lines.append("")
    lines.append("KEY MESSAGE")
    lines.append(content.goal)
    lines.append("")
    lines.append("WALK THROUGH")
    lines.append(f"• {content.context_heading}")
    lines.append(f"  {content.context_body}")
    lines.append("")
    lines.append(f"• {content.benefits_label}")
    for b in content.benefits:
        lines.append(f"  – {b}")
    lines.append("")
    lines.append(content.risk_heading)
    lines.append(content.barrier_intro)
    for item in content.barriers:
        lines.append(f"  – {item.barrier}")
        lines.append(f"    Stakeholders: {item.stakeholders}")
        lines.append(f"    Impact: {item.impact}")
    lines.append("")
    lines.append("MANAGEMENT ASK")
    lines.append(content.management_ask)
    if extra_close:
        lines.append("")
        lines.append(extra_close)
    return "\n".join(lines)


SPEAKER_NOTES: dict[int, str] = {
    1: """OPENING (~30 seconds)

Introduce yourself and the purpose of this session: a status update on the GTM Performance & Readiness portfolio — what we are delivering, where we are blocked, and what we need from leadership.

Set expectations: this is not a deep technical review. It is a programme-level view across five interconnected initiatives, grounded in our requirements register (208 requirements tracked).

Tone: We ARE delivering — but we are sequencing work differently than originally planned because of upstream programme dependencies (One Cisco CRM, Adobe North Star, Business-Ready Dataset).""",

    2: """DOCUMENT OBJECTIVES (~1 minute)

Frame WHY this work matters to the business — before diving into project status.

• Reduce Lead Rejection: Better attribution data at source means the lead scoring model can distinguish genuine buyer intent from passive interest. Poor tagging today contributes to rejected or low-quality leads.

• Boost Sales Trust: Sales needs confidence that leads arriving in SFDC reflect real marketing engagement. Inconsistent tracking erodes that trust.

Transition: "These objectives underpin every programme in this portfolio — UIF, CTT alignment, the Workfront dataset, and ultimately CTT decommission."
""",

    3: """STRATEGIC BENEFITS (~1 minute)

Walk through each bullet — these are the outcomes leadership should expect when the portfolio succeeds:

• Modern analytics standards — moving from Cisco-defined custom IDs to industry-standard UTMs and Workfront IDs.

• One tag, one purpose — every tracking parameter has a clear role; no more mixed CCID usage distorting figures.

• Cross-platform governance — same rules whether activity is created in Workfront, Stensul, or CTT.

• Full touchpoint capture — including non-gated content engagement for top-of-funnel visibility.

• Rich event data — supports both activity (channel) and content attribution.

Transition: "Let me show you where we are across the portfolio."
""",

    4: """PROGRAMME OVERVIEW SECTION (~15 seconds)

"This section covers five programmes. I'll start with UIF — our strategic goal — explain why we pivoted to CTT alignment, then cover the remaining initiatives and what we need from management."
""",

    5: """PORTFOLIO CONTEXT (~2 minutes)

KEY STATS (from requirements register as of August 2026):
• 208 requirements across 5 programmes
• 23 Complete, 19 In Progress/In Review
• Majority still Not Started — concentrated in POC, GTMRP, and Decommission

NARRATIVE:
"We are actively delivering — DTID alignment is fully complete, UIF Phase 1 core tagging is largely done, and CTT FY27 bridge is in progress. But roadblocks from One Cisco CRM, Adobe North Star uncertainty, and no committed capacity for the Business-Ready Dataset force us to keep pivoting."

Emphasise: This is a dependency problem, not a lack of effort or planning.

The five programmes are interconnected — not independent workstreams. Decisions on one (e.g. OCC timeline) affect all others.
""",

    6: _notes_from_summary(
        UIF_SLIDE,
        extra_open="UIF SUMMARY (~3 minutes)\n\nThis is our strategic goal — Phase 1 and Phase 2 combined on one slide.",
        extra_close='TRANSITION: "UIF is the destination — let me explain why we had to pivot to CTT alignment as the FY27 priority."',
    ),

    7: _notes_from_summary(
        UIF_TO_CTT_PIVOT_SLIDE,
        extra_open="PIVOT SLIDE (~2–3 minutes) — EXPECT QUESTIONS HERE\n\nCritical framing: We have NOT abandoned UIF. This is a change of SEQUENCE, not direction.",
        extra_close="""IF PUSHED ON 'WHY NOT STAY ON UIF?':
"Because Reporting needs FY27 data this year, SFDC changes are frozen under OCC, and half our teams aren't on Workfront yet. Waiting for full UIF would mean no FY27 attribution models and broken hybrid reporting."

TRANSITION: "Let me show you what CTT Attribution Data Alignment delivers as the bridge."
""",
    ),

    8: _notes_from_summary(
        CTT_ALIGNMENT_SLIDE,
        extra_open="CTT ATTRIBUTION DATA ALIGNMENT (~3 minutes)\n\nHighlight the win: DTID Reporting Channel is 12/12 Complete — we CAN deliver under constraints.",
        extra_close='TRANSITION: "CTT bridge is in progress. Next — the Workfront–CTT Integration POC, which we have not yet been able to start."',
    ),

    9: """POC SECTION DIVIDER (~10 seconds)

"Workfront & CTT Integration POC — this is the feasibility study to automate ID creation from Workfront into CTT. It has not started yet — I'll explain why that matters."
""",

    10: """POC — DESCRIPTION (~1 minute)

CONTEXT: We don't know how long we'll operate in a dual Workfront/CTT world. Marketing teams currently switch between tools to get Activity and Offer IDs.

The POC asks CETO to prove whether we can auto-generate CCID/OID from Workfront approvals and write them back — removing manual admin.

This is separate from the CTT bridge (which adds FY27 fields). The POC is about process automation and adoption friction.
""",

    11: """POC — GOAL (~45 seconds)

The POC is a DECISION tool — not a build project. Success = evidence to scale, extend the bridge, or pause Workfront expansion.

Thresholds: ≥95% ID creation success, ≥95% FY27 attribute completeness, documented decommission blockers.

Without this evidence, leadership cannot make an informed investment decision on integration vs manual bridge.
""",

    12: """POC — BENEFITS (~1 minute)

Emphasise adoption friction: marketers want to stay in Workfront. Manual CTT ID creation is a top complaint.

POC also protects OCC — no SFDC changes required during pilot.

North Star alignment (BR-09) means integration data should match future Adobe architecture — avoids throwaway work.

Decommission planning: POC documents what blocks CTT retirement.
""",

    13: """POC — BARRIERS (~2 minutes) — ASK FOR CAPACITY

MAIN POINT: 0 of 34 requirements started. No dedicated sprint.

Root cause: Same CTT dev/integration team is fully committed to CTT FY27 bridge (CTT002–CTT018 in progress).

Impact on stakeholders:
• Marketing teams — continued manual ID admin, slower Workfront adoption
• GTMPR — no evidence base for scale/pause/stop decision
• Decommission programme — blockers undocumented

MANAGEMENT ASK: Ring-fence even a 2–3 week POC sprint, separate from bridge maintenance.
""",

    14: """GTMRP SECTION DIVIDER (~10 seconds)

"Business-Ready Workfront Dataset — this is the golden dataset in Snowflake that Reporting, Orchestration, and Attribution teams need. It has not started."
""",

    15: """GTMRP — DESCRIPTION (~1 minute)

Explain the problem: Workfront data today requires complex exports and ad-hoc joins. Every team defines metrics differently.

GTMRP = single business-ready dataset, daily refresh by 8 AM ET, data dictionary, governance framework.

This is the foundation for FY27 attribution models AND seller enrichment — not optional infrastructure.
""",

    16: """GTMRP — GOAL (~45 seconds)

Single source of truth with minimal joins. Self-service for analysts. Governed intake for new fields.

Also resolves the raw vs business-ready architectural decision that is currently blocking CTT bridge field finalisation and UIF Phase 2 alignment.
""",

    17: """GTMRP — BENEFITS (~1 minute)

Connect to leadership priorities:
• Orchestration teams get actionable Workfront insights for sellers
• Reporting gets consistent channel/content attributes at scale
• Resolves circular dependency with CTT bridge — both programmes waiting on architecture decision

Daily refresh + alerting = stakeholders trust the data is current.
""",

    18: """GTMRP — BARRIERS (~2 minutes) — STRONG ASK

MAIN POINT: 0 of 20 requirements. No committed Data Engineering capacity.

This is a programme blocker, not just a project delay:
• Attribution models cannot consume Workfront data at scale
• CTT bridge and GTMRP block each other (raw vs golden schema)
• Leadership cannot see ROI on Workfront investment

MANAGEMENT ASK: Assign named Data Engineering owner and Q1 foundation sprint (GTMRP001–GEN003 minimum).
""",

    19: """DECOMMISSION SECTION DIVIDER (~10 seconds)

"CTT Decommission — leadership directive to move away from Cisco-defined systems. Documented but not scheduled for FY27 execution."
""",

    20: """DECOMMISSION — DESCRIPTION (~1 minute)

49 requirements spanning 15+ systems: Eloqua, SFDC, Tray, Stensul, Demand Intake, Snowflake reporting.

This is the end state — industry-standard Adobe parameters, tracking complexity removed from marketing teams, single Workfront source of truth.

Be honest: decommission BRD exists; integrated cutover plan does not.
""",

    21: """DECOMMISSION — GOAL (~45 seconds)

Retire CTT and legacy CCID/DTID/OID. Replace with Workfront Task IDs and UTMs.

Requires: mandatory Eloqua audit/cleanup, Workfront full adoption, OCC SFDC change window, POC decision.

Phased cutover — not big bang.
""",

    22: """DECOMMISSION — BENEFITS (~1 minute)

Long-term value: reduced dual maintenance, faster lead creation, standardised reporting, North Star alignment.

Contrast with current state: we are EXTENDING CTT for FY27 bridge while decommission BRD assumes retirement — tension leadership needs to resolve.
""",

    23: """DECOMMISSION — BARRIERS (~2 minutes)

KEY MESSAGE: Every other programme is designing around CTT because decommission cannot proceed in FY27.

Blockers:
• OCC — cannot remove CCID from lead creation
• Eloqua audit not started — live OCIDs on web pages and forms
• Workfront adoption incomplete
• POC undecided

MANAGEMENT ASK: Leadership decision — confirm FY27 as intentional bridge year with exit criteria vs commit decommission start in FY28.
""",

    24: """APPENDIX (~10 seconds)

"Supporting documents: Project Summary, requirements register, and generated slides. Happy to share or walk through detail offline."
""",

    25: """CLOSE / Q&A (~1 minute)

SUMMARISE THE THREE ASKS:
1. Acknowledge FY27 as bridge year — UIF goal unchanged, CTT alignment is sequenced delivery
2. Ring-fence capacity — GTMRP foundation sprint + Workfront–CTT POC (even minimal)
3. Provide OCC engagement timeline and lock Reporting/North Star field catalog

Open for questions. If asked about register: 208 reqs, updated as delivery progresses — 44 actively tracked today.

Thank the audience for their time.
""",
}


def get_notes_for_slide(slide_index: int) -> str:
    """Return speaker notes for 1-based slide index."""
    return SPEAKER_NOTES.get(slide_index, "")
