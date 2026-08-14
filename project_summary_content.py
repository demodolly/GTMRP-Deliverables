"""Structured project summary content — Project Summary.docx + management deck."""

from __future__ import annotations

PROJECT_ORDER = [
    "Unified Intelligence Framework (UIF) - Phase 1",
    "Unified Intelligence Framework (UIF) - Phase 2",
    "DTID Channel Alignment / Reporting Channel",
    "CTT & CTT Request Portal FY27 Updates",
    "Workfront & CTT Integration POC",
    "Workfront Business-Ready Dataset (GTMRP)",
    "CTT Decommission",
]

PROJECT_SHORT = {
    "Unified Intelligence Framework (UIF) - Phase 1": "UIF Phase 1",
    "Unified Intelligence Framework (UIF) - Phase 2": "UIF Phase 2",
    "DTID Channel Alignment / Reporting Channel": "DTID Channel Alignment",
    "CTT & CTT Request Portal FY27 Updates": "Workfront to CTT Attribution Data Alignment",
    "Workfront & CTT Integration POC": "Workfront & CTT Integration POC",
    "Workfront Business-Ready Dataset (GTMRP)": "Business-Ready Workfront Dataset",
    "CTT Decommission": "CTT Decommission",
}

# Each project: description, goal, benefits[], barrier_impacts[] as {barrier, impact}
PROJECTS = {
    "Unified Intelligence Framework (UIF) - Phase 1": {
        "description": (
            "Phase 1 of the Unified Intelligence Framework establishes the strategic and technical "
            "foundation to align marketing spend with revenue across Marketo and Eloqua. It enforces "
            "consistent UTM formatting standards across all campaign generation tools (Workfront, "
            "Stensul, manual UTM Builder), captures and preserves utm_id, utm_medium, and utm_source "
            "at every lead touchpoint regardless of marketing automation platform, and deploys dual "
            "ingestion to support legacy CCID/DTID identifiers alongside new UTM schemas for continued "
            "attribution and operational reporting during the transition to Workfront-based tracking."
        ),
        "goal": (
            "Deliver a hybrid-ready tagging and ingestion layer so FY27 reporting, lead attribution, "
            "and operational dashboards continue to function while the organisation transitions from "
            "Drive To ID / CTT identifiers to industry-standard UTM parameters and Workfront IDs — "
            "without breaking existing Salesforce, MSP, or Last Touch reporting."
        ),
        "benefits": [
            "Consistent UTM governance across all online and offline campaign generation tools.",
            "Preserved driving UTMs and legacy CCID/DTID at lead creation — accurate channel and initiative attribution today.",
            "Dual ingestion enables parallel legacy and modern reporting — no big-bang cutover risk.",
            "Standardized Snowflake publishing format — reporting teams can consume one schema.",
            "Foundation for Phase 2 Universal Key (Workfront Content ID) without re-engineering Phase 1 work.",
            "Reduced lead rejection risk through improved attribution data quality at source.",
        ],
        "barrier_impacts": [
            {
                "barrier": "Continued mandatory use of CCID & DTID for SFDC lead creation, MSP, and Last Touch reporting.",
                "impact": "Cannot fully transition to UTM/Channel ID — teams must maintain strict manual governance across Workfront, URL Builder, Adobe, and Ace Reporting; dual maintenance cost persists.",
            },
            {
                "barrier": "Not all marketing teams onboarded to Workfront — mixed operating model.",
                "impact": "Phase 1 benefits only partially realised; non-Workfront teams still depend on CTT IDs and legacy processes, slowing unified attribution.",
            },
            {
                "barrier": "Adobe North Star transition gated by Reporting Team FY27 refresh schedule.",
                "impact": "Future-state visibility delayed — interim Phase 1 work may require adjustment when North Star field model is finalised.",
            },
            {
                "barrier": "Marketo vs Eloqua architecture uncertainty; potential Tealium to Adobe Capture shift (TEA002/TEA003 in review).",
                "impact": "Re-work risk on URL and session capture logic; investment in Marketo path may be wasted if architecture pivots again.",
            },
            {
                "barrier": "GTMPR UTM strategy for extended fields not yet defined (UTM002 deferred).",
                "impact": "Workfront cannot capture extended UTM fields — incomplete strategic data at source for future attribution models.",
            },
            {
                "barrier": "AEM Content ID publishing process not yet operational (PUB002 in progress; Stephen Watts delivery Sep 4).",
                "impact": "Granular content attribution on AEM pages delayed — cannot link customer journey touchpoints to Workfront content assets.",
            },
        ],
    },
    "Unified Intelligence Framework (UIF) - Phase 2": {
        "description": (
            "Phase 2 transitions the organisation to Workfront IDs as the Universal Key — linking "
            "conversations and conversions to strategic marketing data elements for both the initiative "
            "(Channel ID) and the content asset (Content ID). It captures Content IDs on all AEM pages "
            "to monitor every content touchpoint in the customer journey, writes last-touch Content ID "
            "into form submissions in Marketo/Eloqua, and prepares to append strategic Workfront data "
            "elements into SFDC for seller visibility and lead enrichment."
        ),
        "goal": (
            "Achieve 'Zero-Latency Deployment Readiness' — the ability to demonstrate that the moment "
            "a Workfront activation is approved, its Channel ID and Content ID flow automatically through "
            "AEM, MarTech, Tray, and (when permitted) SFDC, enabling true content and initiative "
            "attribution without manual ID administration or legacy CTT dependency."
        ),
        "benefits": [
            "Workfront auto-generates Channel and Content IDs — eliminates manual ID creation for marketing teams.",
            "Hard-linked UTM_ID on every digital URL — conversions tied directly to approved initiatives.",
            "Content ID embedded on every AEM page — full journey attribution to specific content assets.",
            "Identical Universal Key fields across Marketo and Eloqua — consistent ingestion regardless of MAP.",
            "Tray dual-source logic correctly enriches leads from Workfront or legacy IDs.",
            "Hybrid Reporting Mapping document enables YoY analysis without manual translation.",
            "Direct path to Adobe North Star architecture — marketing intelligence captured at source.",
        ],
        "barrier_impacts": [
            {
                "barrier": "One Cisco CRM (OCC) resource constraints — SFDC push intentionally bypassed; no CRM changes permitted in FY27 H1.",
                "impact": "Seller-facing strategic data not visible in SFDC — Phase 2 value limited to backend/logs; sales trust in enriched leads cannot be validated.",
            },
            {
                "barrier": "Reporting teams cannot amend models yet — North Star visibility deferred.",
                "impact": "Cannot prove Phase 2 attribution improvements in production dashboards — business case for full rollout remains unproven.",
            },
            {
                "barrier": "Phase 2 depends on Phase 1 completion (PUB002 AEM Content ID, TEA review items).",
                "impact": "All 12 Phase 2 requirements blocked at Not Started — programme timeline slips with every Phase 1 delay.",
            },
            {
                "barrier": "No synchronized staging environment for dual-ingestion end-to-end validation (DSD003).",
                "impact": "Cannot safely test Workfront + legacy ID coexistence — risk of production data corruption if Phase 2 launches prematurely.",
            },
            {
                "barrier": "Competing priority on CTT FY27 bridge and Phase 1 delivery consumes shared integration capacity.",
                "impact": "Phase 2 receives no dedicated build resource — remains aspirational while bridge work continues.",
            },
        ],
    },
    "DTID Channel Alignment / Reporting Channel": {
        "description": (
            "Because the organisation could not immediately pass UTM values with standardised FY27 "
            "channels, platforms, and vendors, this project created a practical translation layer: "
            "a dedicated Reporting Channel field on DTIDs that maps legacy Vehicle captures to new "
            "FY27 Reporting Channels while preserving old vehicle values until DOPT and Reporting "
            "Teams are ready to rewrite attribution models. It updates the CTT DTID creation, edit, "
            "and search interfaces plus a one-time historical migration."
        ),
        "goal": (
            "Align DTID records with FY27 Channel reporting for attribution models and Adobe data "
            "capture — giving Reporting Teams clean FY27 channel definitions today without forcing "
            "an immediate retirement of legacy DTID values or complex per-report translation logic."
        ),
        "benefits": [
            "Reporting teams can use FY27 channel definitions while DTIDs remain live — no complex translation in every report.",
            "Workfront-first lookup with CTT fallback logic becomes feasible for hybrid activity.",
            "Clean channel selection for new DTIDs — deprecated channels hidden from UI.",
            "Historical DTIDs migrated with audit trail — continuity for FY27 models and YoY analysis.",
            "Data available and ready for when DOPT rewrites attribution models.",
            "All 12 requirements delivered (Complete) — proven delivery capability under programme constraints.",
        ],
        "barrier_impacts": [
            {
                "barrier": "Full DTID audit dependency (audit as of 09 Jun 2026) required before historical migration is final.",
                "impact": "Migration completeness depends on audit quality — gaps could leave historical records misaligned with FY27 channels.",
            },
            {
                "barrier": "Downstream consumers (CTT FY27 bridge, reporting refresh) must adopt Reporting Channel field.",
                "impact": "Value of DTID alignment is not realised until reporting models and CTT bridge reference the new field.",
            },
            {
                "barrier": "Legacy DTIDs and vehicles still active in SFDC during transitional phase.",
                "impact": "Dual identifier world persists — teams must understand both legacy vehicle and new reporting channel semantics.",
            },
        ],
    },
    "CTT & CTT Request Portal FY27 Updates": {
        "description": (
            "Unable to secure Tray or SFDC resources in the first half of FY27 due to the One Cisco "
            "CRM programme, the team pivoted to a bridge approach: add the same strategic FY27 data "
            "elements to CTT Activity IDs and Offer IDs that Workfront captures, so reporting can "
            "use simpler future-proofed logic — first look for a Workfront ID and use corresponding "
            "data elements; if not found, use CTT data with matching attributes. This also supports "
            "CTT decommission planning by monitoring who still uses CTT IDs and migrating them to "
            "Workfront. The CTT Request Portal is updated to match."
        ),
        "goal": (
            "Support FY27 attribution models for channel activations and content consumption during "
            "the hybrid period — ensuring Workfront and non-Workfront teams report on the same FY27 "
            "data elements through CTT/portal bridge fields, without SFDC or Tray changes that would "
            "disrupt One Cisco CRM."
        ),
        "benefits": [
            "FY27 reporting parity between Workfront and non-Workfront teams during hybrid operating model.",
            "Simpler reporting logic — Workfront-first lookup, CTT fallback with identical attributes.",
            "Reduced complex 1:1 mapping logic that does not match cleanly across legacy and modern identifiers.",
            "CTT/Activity ID remains viable for SFDC lead creation while Workfront adoption grows.",
            "Visibility into who still uses CTT IDs — accelerates migration to Workfront.",
            "Portal requestors get FY27 hierarchy, funnel, and stakeholder fields at intake.",
            "OMS and Snowflake receive FY27 attributes for attribution model consumption.",
        ],
        "barrier_impacts": [
            {
                "barrier": "One Cisco CRM freeze — no SFDC or Tray API changes for new CTT fields (CTT016).",
                "impact": "Bridge data cannot flow to CRM in FY27 H1 — reporting must consume from OMS/Snowflake side only; seller enrichment delayed.",
            },
            {
                "barrier": "Not all teams on-boarded to Workfront; CCID still required to create SFDC leads.",
                "impact": "CTT must remain live and enhanced rather than retired — dual tool investment continues.",
            },
            {
                "barrier": "Unclear architectural direction for Workfront data (raw vs business-ready dataset).",
                "impact": "CTT bridge field design may need rework when GTMRP dataset is defined — replanning risk.",
            },
            {
                "barrier": "Portal requirements (POR001–POR010) not yet started — 10 of 26 items still Not Started.",
                "impact": "Non-Workfront requestors still create IDs without FY27 fields at intake — data quality gap at source.",
            },
            {
                "barrier": "Historical Activity/Offer ID audit (CTT006/CTT018) depends on GTMPR audit capacity.",
                "impact": "Cross-time reporting accuracy compromised until historical records are updated.",
            },
            {
                "barrier": "CTT decommission not planned FY27 — bridge scope expands while decommission BRD assumes retirement.",
                "impact": "Competing investment in extending vs retiring CTT — technical debt and team fatigue grow.",
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
    "Workfront Business-Ready Dataset (GTMRP)": {
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
            "Foundation for raw vs business-ready architectural decision — enables CTT bridge and Phase 2 alignment.",
        ],
        "barrier_impacts": [
            {
                "barrier": "No delivery started — 0 of 20 requirements tracked; no committed Data Engineering capacity.",
                "impact": "Attribution models, seller enrichment, and self-service reporting blocked — teams continue ad-hoc joins with inconsistent results.",
            },
            {
                "barrier": "Workfront field definitions still evolving with UIF Phase 1/2 and CTT bridge.",
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
                "barrier": "CTT bridge project waiting on raw vs business-ready architectural decision.",
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
                "impact": "Core decommission step (IT003) impossible in FY27 — CTT cannot be retired while CCID is architecturally required.",
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
