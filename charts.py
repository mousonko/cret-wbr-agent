"""Generate WBR-style charts for the bridge analysis."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import os


def _save(fig, path):
    fig.savefig(path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def chart_compliance_by_site(site_summaries: list, output_dir: str) -> str:
    """Horizontal bar chart: compliance % by site, colored by severity."""
    sites = sorted(site_summaries, key=lambda s: float(s["compliance"].strip("%")) / 100
                   if "%" in str(s["compliance"]) else float(s["compliance"]))

    names = [f"{s['site']} ({s['mp']})" for s in sites]
    vals = []
    for s in sites:
        c = s["compliance"]
        vals.append(float(c.strip("%")) / 100 if "%" in str(c) else float(c))

    colors = ["#d32f2f" if v < 0.5 else "#f57c00" if v < 0.8 else "#388e3c" for v in vals]

    fig, ax = plt.subplots(figsize=(10, max(6, len(names) * 0.35)))
    bars = ax.barh(names, vals, color=colors, edgecolor="white", height=0.7)
    ax.set_xlim(0, 1.05)
    ax.xaxis.set_major_formatter(mticker.PercentFormatter(1.0))
    ax.axvline(x=0.90, color="#388e3c", linestyle="--", alpha=0.7, label="Target 90%")
    ax.set_xlabel("WK-14 Scan Compliance")
    ax.set_title("CRET Scan Compliance by Site — WK-14", fontweight="bold", fontsize=13)
    ax.legend(loc="lower right")

    for bar, val in zip(bars, vals):
        ax.text(val + 0.01, bar.get_y() + bar.get_height() / 2,
                f"{val:.0%}", va="center", fontsize=8)

    plt.tight_layout()
    return _save(fig, os.path.join(output_dir, "chart_compliance_by_site.png"))


def chart_root_cause_pareto(top_causes: list, output_dir: str) -> str:
    """Pareto chart: root causes by number of affected sites."""
    labels = [rc["root_cause"].split("—")[0].strip()[:30] for rc in top_causes]
    counts = [len(rc["affected_sites"]) for rc in top_causes]
    total = sum(counts)
    cumulative = []
    running = 0
    for c in counts:
        running += c
        cumulative.append(running / total)

    fig, ax1 = plt.subplots(figsize=(10, 5))
    bars = ax1.bar(labels, counts, color="#1565c0", edgecolor="white")
    ax1.set_ylabel("# Sites Affected", color="#1565c0")
    ax1.set_title("Root Cause Pareto — WK-14 CRET Scan Compliance", fontweight="bold", fontsize=13)

    ax2 = ax1.twinx()
    ax2.plot(labels, cumulative, color="#d32f2f", marker="o", linewidth=2)
    ax2.yaxis.set_major_formatter(mticker.PercentFormatter(1.0))
    ax2.set_ylabel("Cumulative %", color="#d32f2f")
    ax2.set_ylim(0, 1.1)

    for bar, count in zip(bars, counts):
        ax1.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.2,
                 str(count), ha="center", fontweight="bold")

    plt.xticks(rotation=20, ha="right", fontsize=9)
    plt.tight_layout()
    return _save(fig, os.path.join(output_dir, "chart_root_cause_pareto.png"))


def chart_by_country(site_summaries: list, output_dir: str) -> str:
    """Stacked bar: sites by country and severity."""
    from collections import Counter
    country_sev = {}
    for s in site_summaries:
        mp = s["mp"]
        sev = s["severity"]
        if mp not in country_sev:
            country_sev[mp] = Counter()
        country_sev[mp][sev] += 1

    countries = sorted(country_sev.keys())
    high = [country_sev[c].get("HIGH", 0) for c in countries]
    med = [country_sev[c].get("MEDIUM", 0) for c in countries]
    low = [country_sev[c].get("LOW", 0) for c in countries]

    fig, ax = plt.subplots(figsize=(8, 5))
    ax.bar(countries, high, label="HIGH", color="#d32f2f")
    ax.bar(countries, med, bottom=high, label="MEDIUM", color="#f57c00")
    ax.bar(countries, low, bottom=[h + m for h, m in zip(high, med)], label="LOW", color="#388e3c")
    ax.set_ylabel("# Flagged Sites")
    ax.set_title("Flagged Sites by Country & Severity — WK-14", fontweight="bold", fontsize=13)
    ax.legend()
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    plt.tight_layout()
    return _save(fig, os.path.join(output_dir, "chart_by_country.png"))


def chart_miss_type_breakdown(output_dir: str) -> str:
    """Pie chart of miss types across all sites (from deep-dive data)."""
    # Aggregated from the WK-14 deep-dive data
    labels = ["Dwelling >7 days", "Virtual Depart\n(after FC arrival)", "Depart Scan\nMissing", "Dwelling >7 days\n(extended)"]
    sizes = [246, 60, 121, 128]
    colors = ["#d32f2f", "#f57c00", "#1565c0", "#7b1fa2"]
    explode = (0.05, 0.05, 0.05, 0.05)

    fig, ax = plt.subplots(figsize=(8, 6))
    wedges, texts, autotexts = ax.pie(sizes, explode=explode, labels=labels, colors=colors,
                                       autopct="%1.0f%%", startangle=90, textprops={"fontsize": 10})
    ax.set_title("Miss Type Breakdown — WK-14 (All Sites)", fontweight="bold", fontsize=13)
    plt.tight_layout()
    return _save(fig, os.path.join(output_dir, "chart_miss_type_breakdown.png"))
