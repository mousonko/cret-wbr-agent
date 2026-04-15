"""Generate enhanced WK14 WBR report with charts."""
import json
from charts import (chart_compliance_by_site, chart_root_cause_pareto,
                    chart_by_country, chart_miss_type_breakdown)
from report_generator_v2 import generate_word_enhanced, generate_pptx_enhanced

OUTPUT = "/home/mousonko/.workspace/wbr-bridge-agent/output"

with open(f"{OUTPUT}/wbr_analysis.json") as f:
    analysis = json.load(f)

# Generate charts
print("Generating charts...")
chart_compliance_by_site(analysis["site_summaries"], OUTPUT)
chart_root_cause_pareto(analysis["top_5_root_causes"], OUTPUT)
chart_by_country(analysis["site_summaries"], OUTPUT)
chart_miss_type_breakdown(OUTPUT)

# Generate enhanced reports
print("Generating enhanced Word report...")
generate_word_enhanced(analysis, OUTPUT, f"{OUTPUT}/WBR_Bridge_Analysis_v2.docx")

print("Generating enhanced PowerPoint...")
generate_pptx_enhanced(analysis, OUTPUT, f"{OUTPUT}/WBR_Bridge_Analysis_v2.pptx")

print("\nDone! Enhanced reports:")
print(f"  {OUTPUT}/WBR_Bridge_Analysis_v2.docx")
print(f"  {OUTPUT}/WBR_Bridge_Analysis_v2.pptx")
