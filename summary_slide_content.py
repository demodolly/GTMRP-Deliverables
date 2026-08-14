"""Slide content definitions for single-slide project summaries."""

from summary_slide_builder import BarrierStakeholder, SummarySlideContent

UIF_SLIDE = SummarySlideContent(
    title="Unified Intelligence Framework (UIF): Phase 1 + Phase 2",
    goal=(
        "Move from legacy CCID/DTID to automated Workfront Universal Key attribution across "
        "Marketo & Eloqua — without breaking FY27 reporting, lead creation, or sales operations."
    ),
    context_heading="What We Are Delivering (Phase 1 + Phase 2 Combined)",
    context_body=(
        "Phase 1 (now): UTM governance, dual ingestion of legacy + modern IDs, Snowflake-ready schema.  "
        "Phase 2 (next): Workfront auto-generates Channel & Content IDs, embedded on AEM pages & form "
        "submissions, flowing through MarTech to SFDC when permitted."
    ),
    benefits=[
        "No big-bang cutover — reporting continuity during transition",
        "Accurate attribution at lead creation today",
        "Zero-latency Workfront-to-Marketo/Eloqua automation (Phase 2)",
        "Full content-journey insight & Adobe North Star alignment",
    ],
    barrier_intro=(
        "We are actively delivering UIF capabilities, but upstream dependencies and adoption gaps "
        "limit value realisation for key stakeholder groups:"
    ),
    barriers=[
        BarrierStakeholder(
            "One Cisco CRM freeze — no SFDC changes permitted FY27 H1",
            "Sales, Orchestration, SFDC BCO",
            "Seller-facing Workfront strategic data cannot appear in CRM; Phase 2 value trapped in backend logs; sales cannot validate enriched lead quality.",
        ),
        BarrierStakeholder(
            "CCID & DTID still mandatory for SFDC lead creation, MSP & Last Touch reporting",
            "Marketing, Reporting, GTMPR",
            "Dual manual governance persists; reporting cannot fully adopt UTM/Channel ID; attribution figures at risk of CCID mixing.",
        ),
        BarrierStakeholder(
            "Incomplete Workfront adoption — not all teams onboarded",
            "Marketing, CTT requestors, GTMPR",
            "Non-Workfront teams continue CTT legacy processes; UIF ROI capped; hybrid model persists longer than planned.",
        ),
        BarrierStakeholder(
            "Adobe North Star & Reporting FY27 refresh schedule not locked",
            "Reporting, DOPT, GTMPR",
            "Interim UIF field model may require rework; reporting teams delay model updates; North Star visibility deferred.",
        ),
        BarrierStakeholder(
            "Tealium vs Adobe Capture undecided; Marketo/Eloqua path uncertain",
            "MarTech BCOs, Tealium, Web Analytics",
            "URL/session capture in review; investment in one MAP stack may be wasted; re-work risk on in-flight tagging.",
        ),
        BarrierStakeholder(
            "Business-Ready Workfront Dataset (GTMRP) not started",
            "Reporting, Data Engineering, Orchestration",
            "No single Workfront source of truth; analysts continue ad-hoc joins; UIF field definitions keep shifting.",
        ),
        BarrierStakeholder(
            "Phase 1 incomplete — PUB002 AEM Content ID in progress; 39 reqs Not Started",
            "Publishing, Web, GTMPR, Phase 2 teams",
            "Content-level journey attribution blocked; all Phase 2 requirements cannot start.",
        ),
    ],
    management_ask=(
        "Prioritise OCC engagement window, lock Reporting/North Star field catalog, ring-fence GTMRP "
        "foundation capacity, and complete Phase 1 publishing (PUB002) to unblock Phase 2 stakeholders."
    ),
)

CTT_ALIGNMENT_SLIDE = SummarySlideContent(
    title="CTT Attribution Data Alignment: DTID Channel + Workfront–CTT Bridge",
    goal=(
        "Deliver FY27 reporting parity — clean channel definitions and identical strategic data attributes "
        "whether activity originates in Workfront or CTT — using Workfront-first lookup with CTT fallback, "
        "without SFDC or Tray changes that would disrupt One Cisco CRM."
    ),
    context_heading="What We Are Delivering (DTID Alignment + CTT FY27 Bridge Combined)",
    context_body=(
        "Part 1 — Complete: DTID Reporting Channel field maps legacy Vehicles to FY27 channels (12/12 delivered).  "
        "Part 2 — In Progress: Same FY27 strategic data elements added to CTT Activity IDs, Offer IDs, and the "
        "CTT Request Portal so reporting uses one attribute set across Workfront and non-Workfront activity."
    ),
    benefits=[
        "DTID alignment complete — FY27 channel definitions ready for DOPT model rewrite",
        "Workfront-first / CTT-fallback logic — simpler reporting, no complex 1:1 mapping",
        "Hybrid parity — Workfront and non-Workfront teams report on same FY27 attributes",
        "OMS/Snowflake receive FY27 data for attribution models; CTT usage visible for migration planning",
    ],
    barrier_intro=(
        "DTID delivery proves we can ship under constraints, but the CTT bridge and its stakeholders "
        "remain blocked by upstream programme dependencies:"
    ),
    barriers=[
        BarrierStakeholder(
            "One Cisco CRM freeze — no SFDC or Tray API changes for new CTT bridge fields",
            "Sales, Orchestration, Tray BCO, SFDC BCO",
            "Bridge data cannot reach CRM in FY27 H1; seller enrichment delayed; reporting must consume from OMS/Snowflake only.",
        ),
        BarrierStakeholder(
            "Not all teams on-boarded to Workfront; CCID still required for SFDC lead creation",
            "Marketing, CTT requestors, GTMPR",
            "CTT must remain live and enhanced — dual tool investment continues; non-Workfront teams depend on legacy intake.",
        ),
        BarrierStakeholder(
            "CTT Request Portal not started — POR001–POR010 (10 of 26 items Not Started)",
            "CTT requestors, Marketing ops, GTMPR",
            "Non-Workfront users still create Activity/Offer IDs without FY27 hierarchy, funnel, or stakeholder fields at intake.",
        ),
        BarrierStakeholder(
            "Historical Activity/Offer ID audit (CTT006/CTT018) depends on GTMPR capacity",
            "Reporting, GTMPR, CTT Dev Team",
            "Cross-time reporting accuracy compromised until historical CTT records are updated; YoY analysis at risk.",
        ),
        BarrierStakeholder(
            "DTID historical migration depends on full audit (as of 09 Jun 2026)",
            "Reporting, DOPT, CTT Dev Team",
            "Gaps in DTID audit could leave historical records misaligned with FY27 Reporting Channels.",
        ),
        BarrierStakeholder(
            "Raw vs business-ready Workfront data architecture unresolved (GTMRP not started)",
            "Reporting, Data Engineering, GTMPR",
            "CTT bridge field design may need rework when golden dataset is defined — replanning and duplicate effort.",
        ),
        BarrierStakeholder(
            "CTT decommission not planned FY27 — bridge scope expands vs retirement BRD",
            "Leadership, GTMPR, CTT Dev Team",
            "Competing investment in extending vs retiring CTT; technical debt grows; stakeholder confidence in cutover erodes.",
        ),
    ],
    management_ask=(
        "Confirm FY27 as intentional bridge year with exit criteria; ring-fence portal delivery (POR001–010); "
        "schedule GTMPR historical audit for CTT006/018; resolve GTMRP architecture to lock CTT bridge field design."
    ),
)
