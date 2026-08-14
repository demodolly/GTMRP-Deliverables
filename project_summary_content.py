"""Structured project summary content — Project Summary.docx + management deck."""

from __future__ import annotations

# Five portfolio projects (combined from original BRD/programme split)
PROJECT_ORDER = [
    "Unified Intelligence Framework (UIF)",
    "DTID & CTT Attribution Data Alignment",
    "Workfront & CTT Integration POC",
    "Business-Ready Workfront Dataset (GTMRP)",
    "CTT Decommission",
]

PROJECT_SHORT = {
    "Unified Intelligence Framework (UIF)": "Unified Intelligence Framework (UIF)",
    "DTID & CTT Attribution Data Alignment": "DTID & CTT Attribution Data Alignment",
    "Workfront & CTT Integration POC": "Workfront & CTT Integration POC",
    "Business-Ready Workfront Dataset (GTMRP)": "Business-Ready Workfront Dataset",
    "CTT Decommission": "CTT Decommission",
}

# Register BRD names that roll up into each portfolio project (for stats)
REGISTER_SOURCE_KEYS = {
    "Unified Intelligence Framework (UIF)": [
        "Unified Intelligence Framework (UIF) - Phase 1",
        "Unified Intelligence Framework (UIF) - Phase 2",
    ],
    "DTID & CTT Attribution Data Alignment": [
        "DTID Channel Alignment / Reporting Channel",
        "CTT & CTT Request Portal FY27 Updates",
    ],
    "Workfront & CTT Integration POC": [
        "Workfront & CTT Integration POC",
    ],
    "Business-Ready Workfront Dataset (GTMRP)": [
        "Workfront Business-Ready Dataset (GTMRP)",
    ],
    "CTT Decommission": [
        "CTT Decommission",
    ],
}

PROJECTS = {
    "Unified Intelligence Framework (UIF)": {
        "description": (
            "The Unified Intelligence Framework is a two-phase programme to align marketing spend "
            "with revenue across Marketo and Eloqua. Phase 1 establishes the hybrid foundation: "
            "consistent UTM formatting across all campaign generation tools (Workfront, Stensul, "
            "manual UTM Builder), capture and preservation of utm_id, utm_medium, and utm_source at "
            "every lead touchpoint, and dual ingestion supporting legacy CCID/DTID alongside new "
            "UTM schemas. Phase 2 transitions to Workfront IDs as the Universal Key — auto-generating "
            "Channel and Content IDs, embedding Content IDs on all AEM pages, writing last-touch "
            "Content ID into form submissions, and preparing strategic Workfront data for SFDC when "
            "One Cisco CRM permits."
        ),
        "goal": (
            "Deliver end-to-end marketing intelligence from hybrid-ready tagging (Phase 1) through "
            "Universal Key attribution (Phase 2) — achieving Zero-Latency Deployment Readiness so "
            "that approved Workfront activations flow automatically through AEM, MarTech, Tray, and "
            "SFDC, replacing Drive To ID / CTT dependency with industry-standard parameters without "
            "breaking FY27 reporting, Salesforce lead creation, MSP, or Last Touch reporting."
        ),
        "benefits": [
            "Phase 1: Consistent UTM governance and dual ingestion — legacy and modern reporting in parallel, no big-bang cutover.",
            "Phase 1: Preserved driving UTMs and legacy IDs at lead creation — accurate attribution today.",
            "Phase 1: Standardized Snowflake publishing — one schema for reporting teams.",
            "Phase 2: Workfront auto-generates Channel and Content IDs — eliminates manual ID creation.",
            "Phase 2: Content ID on every AEM page and form submission — full journey and content attribution.",
            "Phase 2: Identical Universal Key fields across Marketo and Eloqua — consistent MAP ingestion.",
            "Combined: Direct path to Adobe North Star — marketing intelligence captured at source.",
            "Combined: Reduced lead rejection through improved attribution data quality.",
        ],
        "barrier_impacts": [
            {
                "barrier": "CCID & DTID still mandatory for SFDC lead creation, MSP, and Last Touch reporting.",
                "impact": "Cannot fully transition to UTM/Channel ID — strict manual governance persists across Workfront, URL Builder, Adobe, and Ace Reporting.",
            },
            {
                "barrier": "Not all marketing teams onboarded to Workfront — mixed operating model.",
                "impact": "UIF benefits only partially realised; non-Workfront teams still depend on CTT IDs and legacy processes.",
            },
            {
                "barrier": "One Cisco CRM freeze — no SFDC changes in FY27 H1; Phase 2 SFDC push bypassed.",
                "impact": "Seller-facing strategic data not visible in SFDC — Phase 2 value limited to backend/logs until OCC permits changes.",
            },
            {
                "barrier": "Adobe North Star and Reporting Team FY27 refresh schedule gate future-state visibility.",
                "impact": "Interim UIF work may require adjustment when North Star field model is finalised — rework risk.",
            },
            {
                "barrier": "Marketo/Eloqua architecture uncertainty; Tealium vs Adobe Capture direction (TEA002/TEA003 in review).",
                "impact": "Re-work risk on URL and session capture; investment in one MAP path may be wasted if architecture pivots.",
            },
            {
                "barrier": "Phase 1 incomplete — PUB002 AEM Content ID in progress; UTM002 deferred; 39 Phase 1 reqs Not Started.",
                "impact": "All 12 Phase 2 requirements blocked — programme timeline slips with every Phase 1 delay.",
            },
            {
                "barrier": "GTMPR UTM strategy for extended fields not yet defined.",
                "impact": "Workfront cannot capture full strategic data at source — incomplete input for Phase 2 attribution models.",
            },
        ],
    },
    "DTID & CTT Attribution Data Alignment": {
        "description": (
            "This combined programme delivers FY27 reporting alignment across legacy and modern "
            "identifier systems in two parts. First, DTID Channel Alignment adds a Reporting Channel "
            "field mapping legacy Vehicle captures to FY27 Reporting Channels while preserving old "
            "values until DOPT rewrites attribution models (12 requirements — Complete). Second, "
            "Workfront to CTT Attribution Data Alignment adds the same strategic FY27 data elements "
            "to CTT Activity IDs, Offer IDs, and the CTT Request Portal — enabling Workfront-first "
            "lookup with CTT fallback when Tray/SFDC resources were unavailable due to One Cisco CRM."
        ),
        "goal": (
            "Give Reporting Teams clean FY27 channel definitions and identical strategic data attributes "
            "whether activity originates in Workfront or CTT — using simpler future-proofed logic "
            "(look for Workfront ID first, fall back to CTT with matching fields) without SFDC or Tray "
            "changes that would disrupt One Cisco CRM, and with data ready for when DOPT rewrites models."
        ),
        "benefits": [
            "DTID Reporting Channel delivered — 12/12 Complete; legacy vehicles mapped to FY27 channels.",
            "Workfront-first / CTT-fallback reporting logic — no complex non-1:1 mapping in every report.",
            "FY27 reporting parity between Workfront and non-Workfront teams during hybrid operating model.",
            "Historical DTIDs migrated — continuity for FY27 models and YoY analysis.",
            "CTT bridge keeps Activity/Offer IDs viable for SFDC lead creation while Workfront adoption grows.",
            "Visibility into who still uses CTT IDs — accelerates migration planning.",
            "OMS and Snowflake receive FY27 attributes for attribution model consumption.",
            "Data available now for when DOPT and Reporting Teams are ready to rewrite models.",
        ],
        "barrier_impacts": [
            {
                "barrier": "One Cisco CRM freeze — no SFDC or Tray API changes for new CTT bridge fields.",
                "impact": "Bridge data cannot flow to CRM in FY27 H1 — reporting must consume from OMS/Snowflake only; seller enrichment delayed.",
            },
            {
                "barrier": "Not all teams on-boarded to Workfront; CCID still required to create SFDC leads.",
                "impact": "CTT must remain live and enhanced — dual tool investment continues instead of retirement.",
            },
            {
                "barrier": "CTT portal requirements (POR001–POR010) not started — 10 of 26 bridge items Not Started.",
                "impact": "Non-Workfront requestors still create IDs without FY27 fields at intake — data quality gap at source.",
            },
            {
                "barrier": "Historical Activity/Offer ID audit (CTT006/CTT018) depends on GTMPR audit capacity.",
                "impact": "Cross-time reporting accuracy compromised until historical CTT records are updated.",
            },
            {
                "barrier": "Full DTID audit dependency (as of 09 Jun 2026) required before migration is final.",
                "impact": "Gaps in DTID audit could leave historical records misaligned with FY27 channels.",
            },
            {
                "barrier": "Unclear raw vs business-ready Workfront data architecture.",
                "impact": "CTT bridge field design may need rework when GTMRP dataset is defined — replanning risk.",
            },
            {
                "barrier": "CTT decommission not planned FY27 — bridge scope expands while retirement BRD assumes cutover.",
                "impact": "Competing investment in extending vs retiring CTT — technical debt grows.",
            },
        ],
    },
    "Workfront & CTT Integration POC": {
        "description": (
            "As the organisation does not know how long it must operate in the dual Workfront/CTT world, "
            "this proof of concept engages CETO to complete a feasibility study on automating generation "
            "of Activity ID (CCID) and Offer ID (OID) from Workfront into CTT. The POC validates whether "
            "integration can remove manual ID administration complexity from marketing teams during "
            "channel and content activation approval workflows."
        ),
        "goal": (
            "Produce decision-quality evidence on whether to scale Workfront–CTT integration, extend "
            "the bridge, or pause Workfront expansion — including time/resource estimates, data quality "
            "thresholds (≥95% ID creation success), and documented CTT decommission blockers."
        ),
        "benefits": [
            "Workfront teams obtain Activity/Offer IDs automatically — no manual CTT ID administration.",
            "Removes adoption friction — marketers stay in Workfront workflow without switching tools.",
            "Auto write-back of generated IDs to Workfront activation records.",
            "Pilot dataset proves FY27 reporting data elements for Workfront-originated activity.",
            "Bridge data aligned to Adobe North Star model — minimises rework.",
            "No SFDC/OCC changes required during POC — protects One Cisco CRM workstream.",
            "Documented decommission blockers and dependencies — clarifies path to CTT retirement.",
        ],
        "barrier_impacts": [
            {
                "barrier": "POC not started — 0 of 34 requirements tracked; no dedicated sprint capacity.",
                "impact": "Marketing teams continue manual ID creation — Workfront adoption friction persists; no evidence for scale/pause/stop decision.",
            },
            {
                "barrier": "Same CTT dev/integration team committed to FY27 bridge delivery (CTT002–CTT018).",
                "impact": "POC competes directly with bridge maintenance — feasibility study timeline unknown.",
            },
            {
                "barrier": "Adobe North Star architecture not finalised (BR-09).",
                "impact": "Integration built in POC may require rework if North Star field model changes.",
            },
            {
                "barrier": "Marketing automation and data lake architecture still unsettled.",
                "impact": "End-to-end POC validation scope limited — cannot prove full MarTech chain.",
            },
            {
                "barrier": "OCC freeze prevents SFDC validation of enriched lead flows.",
                "impact": "POC proves ID automation only — cannot demonstrate seller-facing value.",
            },
        ],
    },
    "Business-Ready Workfront Dataset (GTMRP)": {
        "description": (
            "This project centralises Workfront data elements into a streamlined business-ready golden "
            "dataset in Snowflake to power FY27 Attribution Models and provide Orchestration Teams with "
            "actionable insights for seller enrichment. It includes a metadata repository, data "
            "dictionary, governance framework, and daily-refreshed dataset with channel and content "
            "attributes standardised for self-service analysis."
        ),
        "goal": (
            "Deliver a single-source, business-ready Workfront dataset — available by 8 AM ET daily — "
            "with clear terminology, minimal joins, and governed strategic data elements so Reporting, "
            "Orchestration, and Attribution teams can consume Workfront data confidently without ad-hoc "
            "exports or inconsistent metric logic."
        ),
        "benefits": [
            "Accelerated reporting — analysts use one golden dataset instead of complex Workfront exports.",
            "Consistent metric logic via data dictionary and metadata repository.",
            "FY27 attribution models consume standardised channel and content attributes at scale.",
            "Orchestration teams get actionable seller enrichment insights from Workfront activations.",
            "Governance framework adapts to new data elements while maintaining single source of truth.",
            "Daily refresh with failure alerting — stakeholders trust data currency.",
            "Resolves raw vs business-ready architectural decision — unblocks CTT bridge and UIF Phase 2 alignment.",
        ],
        "barrier_impacts": [
            {
                "barrier": "No delivery started — 0 of 20 requirements tracked; no committed Data Engineering capacity.",
                "impact": "Attribution models, seller enrichment, and self-service reporting blocked — teams continue ad-hoc joins with inconsistent results.",
            },
            {
                "barrier": "Workfront field definitions still evolving with UIF and CTT bridge programmes.",
                "impact": "Dataset design keeps shifting — risk of building wrong schema or repeated rework.",
            },
            {
                "barrier": "Unclear prioritisation vs other FY27 initiatives (CTT bridge, reporting refresh).",
                "impact": "Leadership cannot see Workfront data value — business case for Workfront investment weakens.",
            },
            {
                "barrier": "Adobe North Star and One Cisco CRM roadmaps unsettled — downstream SFDC consumption path unclear.",
                "impact": "Cannot define final seller enrichment fields — GTMRP scope remains ambiguous.",
            },
            {
                "barrier": "CTT bridge programme waiting on raw vs business-ready architectural decision.",
                "impact": "Circular dependency — CTT bridge and GTMRP block each other; programme paralysis on data architecture.",
            },
        ],
    },
    "CTT Decommission": {
        "description": (
            "This programme documents the integrated actions to decommission the Campaign Tagging and "
            "Tracking (CTT) tool — moving to industry-standard and Adobe tracking parameters, removing "
            "tracking complexity from marketing teams through backend automation, and fulfilling a "
            "leadership directive to move away from Cisco-defined systems. It spans 15+ integrated "
            "platforms including Eloqua, SFDC, Tray, Stensul, Demand Intake, and reporting in Snowflake."
        ),
        "goal": (
            "Retire CTT and legacy CCID/DTID/OID identifiers across all marketing, integration, and "
            "reporting systems — replacing them with Workfront Task IDs, UTM parameters, and Universal "
            "Key fields — while maintaining business continuity through a mandatory audit/cleanup phase "
            "and phased cutover aligned to Workfront adoption and One Cisco CRM readiness."
        ),
        "benefits": [
            "Industry-standard Adobe tracking parameters — aligned to North Star architecture.",
            "Tracking complexity removed from marketing teams — automated in backend integrations.",
            "Elimination of Cisco-defined custom systems — reduced maintenance and governance overhead.",
            "Faster, more accurate SFDC lead creation without CCID validation blocks.",
            "Standardized UTM-based reporting — no legacy OCID/CCID filters or CDF dependencies.",
            "Single Workfront source of truth for campaign and content identifiers.",
            "Reduced dual-maintenance cost of running CTT bridge and Workfront in parallel.",
        ],
        "barrier_impacts": [
            {
                "barrier": "49 requirements Not Started — no integrated cutover plan scheduled for FY27.",
                "impact": "Legacy CCID/DTID remain mandatory — every other programme must design around CTT constraints indefinitely.",
            },
            {
                "barrier": "One Cisco CRM blocks SFDC tracking parameter changes required for lead creation without CCID.",
                "impact": "Core decommission step impossible in FY27 — CTT cannot be retired while CCID is architecturally required.",
            },
            {
                "barrier": "Mandatory Eloqua audit/cleanup phase before OCID removal — not yet started.",
                "impact": "High-risk legacy references remain on live web pages, forms, and integrations — decommission would break lead flows.",
            },
            {
                "barrier": "FY27 priority is CTT bridge enhancement, not retirement — competing investment.",
                "impact": "Technical debt grows each quarter; teams lose confidence that decommission will ever happen.",
            },
            {
                "barrier": "Workfront adoption incomplete — non-Workfront teams still require CTT.",
                "impact": "Cannot decommission until all teams migrate — timeline depends on unrelated adoption programme.",
            },
            {
                "barrier": "Workfront–CTT POC undecided — no evidence base for integration vs manual bridge.",
                "impact": "Decommission dependency map incomplete — blockers and exit criteria undocumented.",
            },
        ],
    },
}
