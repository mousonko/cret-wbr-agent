"""Generate WBR report using direct AI analysis of the real WK14 data."""
import json
from report_generator import generate_word, generate_pptx

# AI-analyzed output based on the actual WK14 bridge data
analysis = {
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
    "top_5_root_causes": [
        {
            "rank": 1,
            "root_cause": "Bank Holiday / No B2FC Truck Scheduled — Items dwelling >7 days",
            "affected_sites": ["HBW4", "HBW5", "HBY4", "HRP2", "HBW2", "HNM6", "HSZ2", "HTR2"],
            "impact": "8 DE sites impacted. HBW4 dropped to 11% compliance (16 misses — truck did not arrive 1 Apr, "
                      "BH on 3 Apr prevented rescheduling, items not picked up until 8 Apr). HBY4 at 0% — no truck for 2 weeks. "
                      "This is the single largest driver of WK-14 misses.",
            "recommended_actions": [
                "Schedule pre-BH adhoc B2FC trucks for all DE stations before upcoming holidays",
                "Move HBW4 truck schedule to 00:00 to avoid loading cutoff risk after 07:30",
                "Establish BH contingency SOP: auto-trigger adhoc truck when BH falls on scheduled truck day",
                "HBY4: Review pallet threshold (8 pallets insufficient for additional truck) — adjust trigger"
            ],
            "owner": "DE LM Ops / CRET Transport Planning",
            "timeline": "CW15-16 for immediate recovery; BH SOP by CW17"
        },
        {
            "rank": 2,
            "root_cause": "CRET Backlog Accumulation — Insufficient collection capacity vs volume",
            "affected_sites": ["HBS2", "HSA7", "HIG3"],
            "impact": "3 UK sites with severe backlog. HBS2 at 10% (80 items dwelling >7 days, 14 dwelling >7 days extended). "
                      "HSA7 at 53% (50 items >7 days, 43 extended dwelling). HIG3 at 70% (93 items >7 days). "
                      "Combined 223+ CRET items stuck. CAPs deployed but insufficient collections remain the bottleneck.",
            "recommended_actions": [
                "Source additional adhoc collections for HBS2, HSA7, HIG3 — escalate to transport if needed",
                "Reinstate CAP for HBS2 if collections remain insufficient",
                "HSA7: Assign ownership for depart scan missing (18.2% of misses) — volume-related skip",
                "Target clear floor by WK16; monitor daily until backlog is zero"
            ],
            "owner": "UK LM Ops / Transport Planning",
            "timeline": "WK16-17 for backlog clearance"
        },
        {
            "rank": 3,
            "root_cause": "3PL / SOP Non-Compliance — Missing depart scan at station",
            "affected_sites": ["HBW3", "HNM5", "HSN1"],
            "impact": "HBW3: 10 defects from 3PL not completing final depart scan step. HNM5 at 41% — 27 items (77%) "
                      "with missing depart scan, new SA onboarding. HSN1: SA wrong-scanned items (SA released). "
                      "3PL stations consistently miss the last SOP step.",
            "recommended_actions": [
                "HBW3: 1:1 with 3PL OPS manager completed — verify compliance in CW15",
                "HNM5: Accelerate new SA onboarding; assign buddy for depart scan process",
                "HSN1: Backfill SA role; reinforce depart scan SOP with remaining team",
                "Standardize 3PL CRET depart scan audit — weekly spot check"
            ],
            "owner": "DE LM Ops / 3PL Management",
            "timeline": "CW15-16"
        },
        {
            "rank": 4,
            "root_cause": "Virtual Depart / JEDI Process Errors — Shipments departed without proper scan",
            "affected_sites": ["HHE1", "HBE8", "HCN3", "HBY1", "HLG4"],
            "impact": "HHE1: 17 items (57%) virtual departed after reaching FC — OPs failed JEDI process during backlog reduction. "
                      "HBE8: 11 items (55%) virtual depart. HCN3: 11 items (69%) virtual depart. "
                      "Pattern: stations with backlog rush departures and skip proper scan workflow.",
            "recommended_actions": [
                "HHE1: Backlog now reduced — monitor JEDI compliance from CW15",
                "Investigate HBE8 and HCN3 (no bridge provided) — require bridge by next WBR",
                "Reinforce JEDI process training at stations with >20% virtual depart rate",
                "Add virtual depart as a tracked metric in weekly DS compliance review"
            ],
            "owner": "DE/ES/IT LM Ops",
            "timeline": "CW15 for bridge follow-up; CW16 for process fix"
        },
        {
            "rank": 5,
            "root_cause": "System / Tech Issues — Label and Infinity Line bugs",
            "affected_sites": ["HST1", "HNC1"],
            "impact": "HST1: 48 items (70%) depart scan missing due to new P180 labels not allowing depart — now resolved. "
                      "HNC1: 65 CRETs stuck in Infinity line (50%+ of volume), auto-marked as delivered when DA logged in "
                      "but not actually delivered. Escalated to LME.",
            "recommended_actions": [
                "HST1: Confirm P180 label fix is stable; monitor WK15 depart scan rate",
                "HNC1: Track LME escalation resolution — Infinity line CRET routing fix needed",
                "HST1: New L1 being trained for RTFC processing — verify readiness by CW16",
                "Add tech-related scan failures as a separate category in weekly tracking"
            ],
            "owner": "LME / Tech team + local Ops",
            "timeline": "HST1 stabilizing WK16; HNC1 dependent on LME fix (WK15 target)"
        }
    ],
    "site_summaries": [
        {"site": "HBW4", "mp": "DE", "compliance": "11%", "bridge_summary": "B2FC truck did not arrive 1 Apr + BH Friday prevented rescheduling. Picked up 8 Apr.", "severity": "HIGH"},
        {"site": "HBS2", "mp": "UK", "compliance": "10%", "bridge_summary": "Severe backlog, insufficient collections. CAP deployed but more adhoc collections needed.", "severity": "HIGH"},
        {"site": "HBE8", "mp": "ES", "compliance": "33%", "bridge_summary": "No bridge provided. 55% virtual depart, 45% dwelling >7 days.", "severity": "HIGH"},
        {"site": "HBW5", "mp": "DE", "compliance": "39%", "bridge_summary": "Could not ship B2FC due to bank holidays. Station now clean.", "severity": "HIGH"},
        {"site": "HNM5", "mp": "DE", "compliance": "41%", "bridge_summary": "77% depart scan missing (3PL). New SA onboarding.", "severity": "HIGH"},
        {"site": "HIC1", "mp": "ES", "compliance": "43%", "bridge_summary": "No bridge provided. 100% dwelling >7 days.", "severity": "HIGH"},
        {"site": "HSZ2", "mp": "DE", "compliance": "44%", "bridge_summary": "No bridge provided. 93% dwelling >7 days.", "severity": "HIGH"},
        {"site": "HSA7", "mp": "UK", "compliance": "53%", "bridge_summary": "Ongoing backlog, insufficient collections. Depart scan ownership assigned. Target clear floor WK16.", "severity": "HIGH"},
        {"site": "HST1", "mp": "UK", "compliance": "59%", "bridge_summary": "P180 label tech issue preventing depart scan — now resolved. New L1 training.", "severity": "MEDIUM"},
        {"site": "HHE1", "mp": "DE", "compliance": "60%", "bridge_summary": "Backlog reduction caused JEDI process errors. Volume now under control.", "severity": "MEDIUM"},
        {"site": "HSN1", "mp": "DE", "compliance": "65%", "bridge_summary": "SA wrong-scanned items. SA released.", "severity": "MEDIUM"},
        {"site": "HNC1", "mp": "FR", "compliance": "66%", "bridge_summary": "CRETs stuck in Infinity line — auto-marked delivered incorrectly. Escalated to LME.", "severity": "MEDIUM"},
        {"site": "HLG4", "mp": "IT", "compliance": "70%", "bridge_summary": "No bridge provided. Mix of dwelling, missing depart scan, and virtual depart.", "severity": "MEDIUM"},
        {"site": "HIG3", "mp": "UK", "compliance": "70%", "bridge_summary": "Backlog from high collections volume. Adhoc collections needed, CAP may be required.", "severity": "MEDIUM"},
        {"site": "HBW2", "mp": "DE", "compliance": "73%", "bridge_summary": "No bridge provided. 86% dwelling >7 days.", "severity": "MEDIUM"},
        {"site": "HNM6", "mp": "DE", "compliance": "78%", "bridge_summary": "No bridge provided. 93% dwelling >7 days.", "severity": "MEDIUM"},
        {"site": "HCN3", "mp": "ES", "compliance": "79%", "bridge_summary": "No bridge provided. 69% virtual depart.", "severity": "MEDIUM"},
        {"site": "HTR2", "mp": "DE", "compliance": "81%", "bridge_summary": "No bridge provided. 91% dwelling >7 days.", "severity": "LOW"},
        {"site": "HRP2", "mp": "DE", "compliance": "84%", "bridge_summary": "BH period — no adhoc possible. 1 wrong stow retrained.", "severity": "LOW"},
        {"site": "HBW3", "mp": "DE", "compliance": "85%", "bridge_summary": "10 defects from 3PL missing depart scan. 1:1 with OPS manager done.", "severity": "LOW"},
        {"site": "HVQ3", "mp": "ES", "compliance": "88%", "bridge_summary": "No bridge provided. Minor: 3 dwelling + 1 depart scan miss.", "severity": "LOW"},
        {"site": "HBY1", "mp": "DE", "compliance": "90%", "bridge_summary": "5 misses — 3 from BH 2 Apr, 2 from DA next-day return. DSP informed.", "severity": "LOW"},
        {"site": "HBY4", "mp": "DE", "compliance": "0%", "bridge_summary": "No truck for 2 weeks (last truck 27 Mar). 8 pallets insufficient for adhoc trigger.", "severity": "HIGH"},
    ],
    "patterns_and_trends": (
        "1. DE dominates (13/23 sites) — almost entirely driven by Easter bank holiday truck scheduling gaps. "
        "Most DE sites will self-recover by CW16 once trucks resume.\n"
        "2. UK has a structural backlog problem (HBS2, HSA7, HIG3) — not holiday-related. "
        "Collection capacity is consistently below CRET volume. CAPs help temporarily but don't solve the root cause.\n"
        "3. ES has 4 flagged sites but 0 bridges provided — accountability gap. Need POC assignment and bridge enforcement.\n"
        "4. Virtual depart is an emerging pattern (HHE1, HBE8, HCN3, HBY1, HLG4) — stations under pressure "
        "skip proper JEDI scan workflow. This masks the real compliance picture.\n"
        "5. 3PL stations (HBW3, HNM5) consistently miss the depart scan step — SOP adherence is a recurring theme."
    ),
    "recommended_wbr_talking_points": [
        "23 sites flagged in WK-14 (up from WK-13 baseline) — 9 without bridges, primarily ES. Action: enforce bridge completion by WK15 WBR.",
        "Easter BH is the #1 driver: 8 DE sites impacted by no B2FC trucks. Most will self-recover CW15-16. Preventive action: pre-BH truck scheduling SOP needed before next holiday period.",
        "UK backlog is structural, not seasonal: HBS2 (10%), HSA7 (53%), HIG3 (70%) need additional collection capacity. Escalate to transport if adhoc sourcing fails by WK16.",
        "3PL depart scan compliance remains a gap: HBW3 and HNM5 — recommend weekly 3PL audit cadence.",
        "Tech issues at HST1 (P180 labels) and HNC1 (Infinity line) — both escalated, monitor resolution WK15."
    ]
}

# Generate reports
generate_word(analysis, "/home/mousonko/.workspace/wbr-bridge-agent/output/WBR_Bridge_Analysis.docx")
generate_pptx(analysis, "/home/mousonko/.workspace/wbr-bridge-agent/output/WBR_Bridge_Analysis.pptx")

with open("/home/mousonko/.workspace/wbr-bridge-agent/output/wbr_analysis.json", "w") as f:
    json.dump(analysis, f, indent=2, ensure_ascii=False)

print("Reports generated successfully!")
print("  - output/WBR_Bridge_Analysis.docx")
print("  - output/WBR_Bridge_Analysis.pptx")
print("  - output/wbr_analysis.json")
