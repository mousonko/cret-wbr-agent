"""Generate WK14 report with narrative-style root cause bridges."""
import json
import os
from report_generator_v3 import generate_narrative_word, generate_narrative_pptx

OUTPUT = "/home/mousonko/.workspace/wbr-bridge-agent/output"

# Load WoW data if available
wow = {}
wow_path = os.path.join(OUTPUT, "wow_analysis.json")
if os.path.exists(wow_path):
    with open(wow_path) as f:
        wow = json.load(f)

analysis = {
    "week": "Week 14",
    "executive_summary": (
        "23 sites flagged for CRET scan compliance in WK-14 across DE (13), ES (4), UK (4), FR (1), and IT (1). "
        "Only 14 of 23 sites provided bridges — 9 sites (mostly ES) have no bridge or POC assigned, requiring follow-up. "
        "EU-wide compliance dropped significantly vs T4W trending, driven primarily by Easter/bank holiday disruptions "
        "(no B2FC trucks scheduled), CRET backlog accumulation in UK stations, and process non-compliance at 3PL stations. "
        "Worst performers: HBW4 (11%), HBS2 (10%), HBE8 (33%), HBW5 (39%), HIC1 (43%)."
    ),
    "total_flagged_sites": 23,
    "sites_with_bridges": 14,
    "sites_without_bridges": 9,

    # === NEW NARRATIVE FORMAT ===
    "bridge_summary": (
        "23 sites flagged in WK-14 (up from WK-13 baseline) — 9 without bridges, primarily ES. "
        "Action: enforce bridge completion by CWK16 CRET WBR. Owner: ES MSL."
    ),
    "top_root_causes_narrative": [
        {
            "rank": 1,
            "title": "Easter BH is the #1 driver:",
            "narrative": (
                "8 DE sites impacted by no B2FC trucks. Most will self-recover CW15-16. "
                "Preventive action 1: sites to raise ad-hoc truck request prior BH to avoid delay and miss pick up opportunities. "
                "Preventive action 2: pre-BH truck scheduling SOP needed before next holiday period aligned with Feran Ferrer. "
                "ETA: WK19. Owner: mousonko."
            ),
        },
        {
            "rank": 2,
            "title": "UK backlog issue -> structural topic, not seasonal:",
            "narrative": (
                "HBS2 (10%), HSA7 (53%), HIG3 (70%) need additional collection capacity. "
                "Escalate to transport and align with Relo sites if adhoc sourcing fails by EOW16. "
                "Owner: HSA7/HIG3 MSL."
            ),
        },
        {
            "rank": 3,
            "title": "SOP adherence in some 3PL, depart scan compliance remains a gap:",
            "narrative": (
                "HBW3 and HNM5 — recommend weekly 3PL audit cadence. "
                "SAM to ensure SOP compliance by providing refresher training to 3PL. ETA: EOW16."
            ),
        },
        {
            "rank": 4,
            "title": "Tech issues at HST1 (P180 labels) and HNC1 (Infinity line):",
            "narrative": (
                "Both escalated and resolved. SAM to monitor resolution WK16."
            ),
        },
    ],

    # Keep structured data for charts and tables
    "top_5_root_causes": [
        {"rank": 1, "root_cause": "Bank Holiday / No B2FC Truck", "affected_sites": ["HBW4", "HBW5", "HBY4", "HRP2", "HBW2", "HNM6", "HSZ2", "HTR2"],
         "impact": "8 DE sites", "recommended_actions": ["Pre-BH adhoc truck requests", "BH truck scheduling SOP"], "owner": "mousonko", "timeline": "WK19"},
        {"rank": 2, "root_cause": "UK Backlog — Insufficient collections", "affected_sites": ["HBS2", "HSA7", "HIG3"],
         "impact": "3 UK sites", "recommended_actions": ["Escalate to transport", "Align with Relo sites"], "owner": "HSA7/HIG3 MSL", "timeline": "EOW16"},
        {"rank": 3, "root_cause": "3PL SOP Non-Compliance", "affected_sites": ["HBW3", "HNM5"],
         "impact": "2 DE sites", "recommended_actions": ["Weekly 3PL audit", "SAM refresher training"], "owner": "SAM", "timeline": "EOW16"},
        {"rank": 4, "root_cause": "System / Tech Issues", "affected_sites": ["HST1", "HNC1"],
         "impact": "2 sites", "recommended_actions": ["Monitor resolution WK16"], "owner": "SAM", "timeline": "WK16"},
    ],
    "site_summaries": [
        {"site": "HBW4", "mp": "DE", "compliance": "11%", "bridge_summary": "B2FC truck did not arrive 1 Apr + BH Friday prevented rescheduling.", "severity": "HIGH"},
        {"site": "HBS2", "mp": "UK", "compliance": "10%", "bridge_summary": "Severe backlog, insufficient collections. CAP deployed.", "severity": "HIGH"},
        {"site": "HBE8", "mp": "ES", "compliance": "33%", "bridge_summary": "No bridge provided. 55% virtual depart, 45% dwelling >7 days.", "severity": "HIGH"},
        {"site": "HBW5", "mp": "DE", "compliance": "39%", "bridge_summary": "Could not ship B2FC due to bank holidays. Station now clean.", "severity": "HIGH"},
        {"site": "HNM5", "mp": "DE", "compliance": "41%", "bridge_summary": "77% depart scan missing (3PL). New SA onboarding.", "severity": "HIGH"},
        {"site": "HIC1", "mp": "ES", "compliance": "43%", "bridge_summary": "No bridge provided. 100% dwelling >7 days.", "severity": "HIGH"},
        {"site": "HSZ2", "mp": "DE", "compliance": "44%", "bridge_summary": "No bridge provided. 93% dwelling >7 days.", "severity": "HIGH"},
        {"site": "HSA7", "mp": "UK", "compliance": "53%", "bridge_summary": "Ongoing backlog, insufficient collections. Target clear floor WK16.", "severity": "HIGH"},
        {"site": "HST1", "mp": "UK", "compliance": "59%", "bridge_summary": "P180 label tech issue — now resolved. New L1 training.", "severity": "MEDIUM"},
        {"site": "HHE1", "mp": "DE", "compliance": "60%", "bridge_summary": "Backlog reduction caused JEDI process errors. Volume now under control.", "severity": "MEDIUM"},
        {"site": "HSN1", "mp": "DE", "compliance": "65%", "bridge_summary": "SA wrong-scanned items. SA released.", "severity": "MEDIUM"},
        {"site": "HNC1", "mp": "FR", "compliance": "66%", "bridge_summary": "CRETs stuck in Infinity line — auto-marked delivered incorrectly.", "severity": "MEDIUM"},
        {"site": "HLG4", "mp": "IT", "compliance": "70%", "bridge_summary": "No bridge provided. Mix of dwelling, missing depart scan.", "severity": "MEDIUM"},
        {"site": "HIG3", "mp": "UK", "compliance": "70%", "bridge_summary": "Backlog from high collections volume. Adhoc collections needed.", "severity": "MEDIUM"},
        {"site": "HBW2", "mp": "DE", "compliance": "73%", "bridge_summary": "No bridge provided. 86% dwelling >7 days.", "severity": "MEDIUM"},
        {"site": "HNM6", "mp": "DE", "compliance": "78%", "bridge_summary": "No bridge provided. 93% dwelling >7 days.", "severity": "MEDIUM"},
        {"site": "HCN3", "mp": "ES", "compliance": "79%", "bridge_summary": "No bridge provided. 69% virtual depart.", "severity": "MEDIUM"},
        {"site": "HTR2", "mp": "DE", "compliance": "81%", "bridge_summary": "No bridge provided. 91% dwelling >7 days.", "severity": "LOW"},
        {"site": "HRP2", "mp": "DE", "compliance": "84%", "bridge_summary": "BH period — no adhoc possible. 1 wrong stow retrained.", "severity": "LOW"},
        {"site": "HBW3", "mp": "DE", "compliance": "85%", "bridge_summary": "10 defects from 3PL missing depart scan. 1:1 with OPS manager done.", "severity": "LOW"},
        {"site": "HVQ3", "mp": "ES", "compliance": "88%", "bridge_summary": "No bridge provided. Minor: 3 dwelling + 1 depart scan miss.", "severity": "LOW"},
        {"site": "HBY1", "mp": "DE", "compliance": "90%", "bridge_summary": "5 misses — 3 from BH 2 Apr, 2 from DA next-day return.", "severity": "LOW"},
        {"site": "HBY4", "mp": "DE", "compliance": "0%", "bridge_summary": "No truck for 2 weeks. 8 pallets insufficient for adhoc trigger.", "severity": "HIGH"},
    ],
    "recommended_wbr_talking_points": [
        "23 sites flagged in WK-14 — 9 without bridges, primarily ES. Action: enforce bridge completion by CWK16.",
        "Easter BH is the #1 driver: 8 DE sites. Pre-BH truck scheduling SOP needed. ETA: WK19.",
        "UK backlog is structural: HBS2 (10%), HSA7 (53%), HIG3 (70%). Escalate to transport by EOW16.",
        "3PL depart scan compliance gap: HBW3, HNM5. SAM refresher training by EOW16.",
        "Tech issues at HST1 and HNC1 — both resolved. Monitor WK16.",
    ],
}

# Save updated analysis
with open(os.path.join(OUTPUT, "wbr_analysis.json"), "w") as f:
    json.dump(analysis, f, indent=2, ensure_ascii=False)

# Generate reports
generate_narrative_word(analysis, wow, OUTPUT)
generate_narrative_pptx(analysis, wow, OUTPUT)
print("Done!")
