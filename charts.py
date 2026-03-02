"""
charts.py — Generate interactive Plotly HTML charts for AI tool adoption trends.

Per tool per function:
  - AI Code / NABU  → one chart with Adoption Rate + Active Rate (two lines)
  - Microsoft Copilot → same rate chart, plus a dual-axis chart for
    Total Actions (left Y) and Avg Action Per User (right Y)
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd
import plotly.graph_objects as go

# ---------------------------------------------------------------------------
# Colour palette
# ---------------------------------------------------------------------------
TOOL_COLORS: dict[str, str] = {
    "AI Code": "#6C9EFF",
    "NABU": "#FF6B6B",
    "Microsoft Copilot": "#4ECDC4",
}
_EXTRA_COLORS = ["#C084FC", "#FDBA74", "#38BDF8", "#FB7185", "#BEF264"]

CHART_HEIGHT = 310

_DARK_LAYOUT: dict[str, Any] = dict(
    template="plotly_dark",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="#1e2433",
    font=dict(color="#c8cdd5", size=11),
)

RATE_COLORS = {"Adoption Rate": "#6EE7B7", "Active Rate": "#FBBF24"}
ACTION_COLORS = {"Total Actions": "#93C5FD", "Avg Action Per User": "#F9A8D4"}


def _color_for_tool(tool: str) -> str:
    if tool in TOOL_COLORS:
        return TOOL_COLORS[tool]
    return _EXTRA_COLORS[hash(tool) % len(_EXTRA_COLORS)]


def _hex_to_rgba(hex_color: str, alpha: float) -> str:
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


# ---------------------------------------------------------------------------
# Month sorting
# ---------------------------------------------------------------------------

def _month_sort_key(label: str) -> tuple[int, int]:
    month_map = {
        "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
        "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
    }
    m = re.match(r"[A-Za-z]{3,}\.\s*\d{2}\s*-\s*([A-Za-z]{3,})\.\s*(\d{2})", label)
    if m:
        return (int(m.group(2)) + 2000, month_map.get(m.group(1).lower()[:3], 0))
    m = re.match(r"([A-Za-z]{3,})\.\s*(\d{2})", label)
    if m:
        return (int(m.group(2)) + 2000, month_map.get(m.group(1).lower()[:3], 0))
    return (0, 0)


def _sorted_months(months) -> list[str]:
    return sorted(set(months), key=_month_sort_key)


_MONTH_ABBRS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _full_year_months(data_months: list[str]) -> list[str]:
    """Extend data months to a full 12-month calendar year.

    Range-format months (e.g. 'Dec. 24 - Nov. 25') are prepended before
    the 12 calendar months of the latest year found in the data.
    """
    sorted_data = _sorted_months(data_months)

    # Separate range months from single months
    range_months = [m for m in sorted_data if "-" in m]
    single_months = [m for m in sorted_data if "-" not in m]

    if not single_months:
        return sorted_data

    # Find the latest year from single months
    years = set()
    for m in single_months:
        yr, _ = _month_sort_key(m)
        years.add(yr)
    latest_year = max(years)
    yr_short = latest_year - 2000

    # Generate full 12 months for that year
    full_calendar = [f"{abbr}. {yr_short}" for abbr in _MONTH_ABBRS]

    return range_months + full_calendar


# ---------------------------------------------------------------------------
# Chart builders
# ---------------------------------------------------------------------------

def _make_rate_chart(
    df: pd.DataFrame,
    tool: str,
    function_name: str,
    all_months: list[str] | None = None,
) -> go.Figure | None:
    """Adoption Rate + Active Rate on one chart for a single tool/function."""
    sub = df[(df["Tool"] == tool) & (df["Function"] == function_name)].copy()
    if sub.empty:
        return None

    months = all_months if all_months else _sorted_months(sub["Month"].tolist())
    sub = sub.drop_duplicates(subset="Month", keep="first")
    sub = sub.set_index("Month").reindex(months).reset_index()

    fig = go.Figure()
    x_vals = sub["Month"].tolist()
    for metric, color in RATE_COLORS.items():
        y = pd.to_numeric(sub[metric], errors="coerce")
        y_list = y.tolist()
        fig.add_trace(go.Scatter(
            x=x_vals, y=y_list,
            mode="lines+markers+text",
            name=metric,
            line=dict(color=color, width=2.5),
            marker=dict(size=7, color=color),
            text=[f"{v:.1%}" if pd.notna(v) else "" for v in y_list],
            textposition="top center",
            textfont=dict(size=10, color=color),
            connectgaps=True,
        ))

    fig.update_layout(
        title=dict(text=tool, x=0.5, font=dict(size=15)),
        yaxis=dict(tickformat=".0%", gridcolor="rgba(255,255,255,0.06)",
                   title=None, range=[0, 1]),
        xaxis=dict(gridcolor="rgba(255,255,255,0.06)", title=None, type="category"),
        height=CHART_HEIGHT,
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="center", x=0.5, font=dict(size=10)),
        margin=dict(t=60, b=36, l=48, r=16),
        **_DARK_LAYOUT,
    )
    return fig


def _make_actions_chart(
    df: pd.DataFrame,
    function_name: str,
    all_months: list[str] | None = None,
) -> go.Figure | None:
    """Dual-axis chart: Total Actions (left Y, bar) + Avg Action/User (right Y, line).
    Only for Microsoft Copilot.
    """
    sub = df[(df["Tool"] == "Microsoft Copilot") & (df["Function"] == function_name)].copy()
    if sub.empty:
        return None
    if "Total Actions" not in sub.columns or "Avg Action Per User" not in sub.columns:
        return None

    total_acts = pd.to_numeric(sub["Total Actions"], errors="coerce")
    avg_acts = pd.to_numeric(sub["Avg Action Per User"], errors="coerce")
    if total_acts.isna().all() and avg_acts.isna().all():
        return None

    months = all_months if all_months else _sorted_months(sub["Month"].tolist())
    sub = sub.drop_duplicates(subset="Month", keep="first")
    sub = sub.set_index("Month").reindex(months).reset_index()

    total_acts = pd.to_numeric(sub["Total Actions"], errors="coerce").tolist()
    avg_acts = pd.to_numeric(sub["Avg Action Per User"], errors="coerce").tolist()
    x_vals = sub["Month"].tolist()

    fig = go.Figure()

    c1 = ACTION_COLORS["Total Actions"]
    fig.add_trace(go.Bar(
        x=x_vals, y=total_acts,
        name="Total Actions",
        marker_color=_hex_to_rgba(c1, 0.75),
        text=[f"{v:,.0f}" if pd.notna(v) else "" for v in total_acts],
        textposition="outside",
        textfont=dict(size=9, color=c1),
        offsetgroup="total_actions",
        yaxis="y",
    ))

    c2 = ACTION_COLORS["Avg Action Per User"]
    fig.add_trace(go.Bar(
        x=x_vals, y=avg_acts,
        name="Avg Action/User",
        marker_color=_hex_to_rgba(c2, 0.75),
        text=[f"{v:.1f}" if pd.notna(v) else "" for v in avg_acts],
        textposition="outside",
        textfont=dict(size=10, color=c2),
        offsetgroup="avg_actions",
        yaxis="y2",
    ))

    fig.update_layout(
        title=dict(text="Copilot Actions", x=0.5, font=dict(size=15)),
        yaxis=dict(
            title=dict(text="Total Actions", font=dict(size=10, color=c1)),
            gridcolor="rgba(255,255,255,0.06)",
            tickfont=dict(color=c1),
        ),
        yaxis2=dict(
            title=dict(text="Avg Action/User", font=dict(size=10, color=c2)),
            overlaying="y", side="right",
            gridcolor="rgba(255,255,255,0.0)",
            tickfont=dict(color=c2),
            range=[0, 1000],
        ),
        xaxis=dict(gridcolor="rgba(255,255,255,0.06)", title=None, type="category"),
        height=CHART_HEIGHT,
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="center", x=0.5, font=dict(size=10)),
        margin=dict(t=60, b=36, l=56, r=56),
        barmode="group",
        bargap=0.25,
        **_DARK_LAYOUT,
    )
    return fig


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------
ChartEntry = dict  # keys: tool, kind, html
FunctionSection = tuple[str, str, list[ChartEntry]]  # (label, anchor, charts)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate_charts(df: pd.DataFrame, output_dir: str | Path) -> Path:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    for old_html in output_dir.glob("*.html"):
        if old_html.name != "dashboard.html":
            old_html.unlink(missing_ok=True)

    functions = sorted([f for f in df["Function"].unique() if f and f != "Total"])
    groups = ["Total"] + functions
    all_tools = sorted(df["Tool"].unique())
    months = _sorted_months(df["Month"].tolist())
    full_months = _full_year_months(df["Month"].tolist())
    has_actions = ("Total Actions" in df.columns and "Avg Action Per User" in df.columns)

    sections: list[FunctionSection] = []
    chart_count = 0

    for group in groups:
        label = "Total (All Functions)" if group == "Total" else group
        safe = group.replace(" ", "_").replace("/", "_")
        anchor = safe.lower()
        entries: list[ChartEntry] = []

        # --- One rate chart per tool ---
        for tool in all_tools:
            fig = _make_rate_chart(df, tool, group, all_months=full_months)
            if fig:
                html = fig.to_html(full_html=False, include_plotlyjs=False)
                entries.append({"tool": tool, "kind": "rates", "html": html})
                chart_count += 1

        # --- Copilot actions chart ---
        if has_actions:
            fig = _make_actions_chart(df, group, all_months=full_months)
            if fig:
                html = fig.to_html(full_html=False, include_plotlyjs=False)
                entries.append({"tool": "Microsoft Copilot", "kind": "actions", "html": html})
                chart_count += 1

        if entries:
            sections.append((label, anchor, entries))

    dashboard = output_dir / "dashboard.html"
    _write_dashboard(dashboard, sections, months, all_tools)
    print(f"\n  Embedded {chart_count} charts into 1 dashboard")
    print(f"  Output directory: {output_dir}")
    return dashboard


# ---------------------------------------------------------------------------
# Dashboard HTML
# ---------------------------------------------------------------------------

def _write_dashboard(
    path: Path,
    sections: list[FunctionSection],
    months: list[str],
    tools: list[str],
) -> None:
    toc = "\n".join(
        f'      <li><a href="#{a}">{l}</a></li>' for l, a, _ in sections
    )

    body_parts: list[str] = []
    for label, anchor, entries in sections:
        rate_cards = [e for e in entries if e["kind"] == "rates"]
        action_cards = [e for e in entries if e["kind"] == "actions"]

        # Split rate cards: non-Copilot vs Copilot
        non_copilot_rates = [e for e in rate_cards if e["tool"] != "Microsoft Copilot"]
        copilot_rate = [e for e in rate_cards if e["tool"] == "Microsoft Copilot"]

        body_parts.append(f'<section id="{anchor}" class="section">')
        body_parts.append(f'  <h2 class="section-hdr">{label}</h2>')

        # --- Row 1: AI Code + NABU rate charts ---
        if non_copilot_rates:
            n = len(non_copilot_rates)
            body_parts.append(f'  <h3 class="metric-label">Adoption Rate &amp; Active Rate</h3>')
            body_parts.append(f'  <div class="card-row" style="grid-template-columns:repeat({n},1fr)">')
            for e in non_copilot_rates:
                body_parts.append(_card_html(e))
            body_parts.append('  </div>')

        # --- Row 2: Copilot rate chart + Copilot actions chart ---
        if copilot_rate or action_cards:
            copilot_entries = copilot_rate + action_cards
            n = len(copilot_entries)
            body_parts.append(f'  <h3 class="metric-label">Microsoft Copilot — Rates &amp; Actions</h3>')
            body_parts.append(f'  <div class="card-row" style="grid-template-columns:repeat({n},1fr)">')
            for e in copilot_entries:
                body_parts.append(_card_html(e))
            body_parts.append('  </div>')

        body_parts.append('</section>')

    charts_body = "\n".join(body_parts)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>AI Tool Adoption Dashboard</title>
<script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
<style>
  *, *::before, *::after {{ box-sizing: border-box; }}
  body {{
    font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
    margin: 0; padding: 20px 28px 48px;
    background: #0f1117; color: #c8cdd5; line-height: 1.5;
  }}

  /* header */
  .header {{
    max-width: 1500px; margin: 0 auto 24px;
    padding: 24px 28px 18px;
    background: linear-gradient(135deg, #161b26 0%, #1a2236 100%);
    border: 1px solid #262d3d; border-radius: 12px; color: #e4e8ef;
  }}
  .header h1 {{ margin: 0 0 6px; font-size: 1.35rem; font-weight: 600; letter-spacing: -.02em; }}
  .header .meta {{ color: rgba(200,210,225,.6); font-size: .76rem; margin: 0; }}
  .header .meta span {{
    display: inline-block; background: rgba(255,255,255,.07);
    padding: 2px 10px; border-radius: 4px; margin-right: 6px;
  }}

  /* toc */
  .toc {{
    max-width: 1500px; margin: 0 auto 24px;
    background: #161b26; border: 1px solid #262d3d;
    padding: 16px 24px; border-radius: 10px;
  }}
  .toc h2 {{ margin: 0 0 8px; font-size: .88rem; color: #8b95a5; }}
  .toc ul {{ margin: 0; padding-left: 16px; columns: 3; column-gap: 28px; }}
  .toc li {{ margin: 3px 0; break-inside: avoid; font-size: .76rem; }}
  .toc a {{ color: #7aa2f7; text-decoration: none; }}
  .toc a:hover {{ text-decoration: underline; }}

  /* sections */
  .section {{ max-width: 1500px; margin: 0 auto 32px; }}
  .section-hdr {{
    font-size: 1rem; font-weight: 600; color: #e4e8ef;
    margin: 0 0 12px; padding-bottom: 6px;
    border-bottom: 1px solid #2a3144;
  }}
  .metric-label {{
    font-size: .72rem; font-weight: 600; text-transform: uppercase;
    letter-spacing: .06em; color: #6b7280; margin: 0 0 8px;
  }}

  /* card grid */
  .card-row {{ display: grid; gap: 12px; margin-bottom: 18px; }}
  @media (max-width: 900px) {{ .card-row {{ grid-template-columns: 1fr !important; }} }}

  /* card */
  .card {{
    background: #161b26; border: 1px solid #262d3d;
    border-radius: 10px; overflow: hidden;
    transition: border-color .15s, box-shadow .15s;
  }}
  .card:hover {{ border-color: #3a4560; box-shadow: 0 4px 20px rgba(0,0,0,.35); }}
  .card-hdr {{
    display: flex; justify-content: space-between; align-items: center;
    padding: 8px 14px 4px;
  }}
  .tool-badge {{
    font-size: .62rem; font-weight: 700; color: #fff;
    padding: 2px 10px; border-radius: 4px; letter-spacing: .03em;
  }}
  .ext {{ color: #7aa2f7; text-decoration: none; font-size: .7rem; }}
  .ext:hover {{ text-decoration: underline; }}
  .card-chart {{ padding: 0 4px 4px; }}
  .card-chart .plotly-graph-div {{
    width: 100% !important;
    height: {CHART_HEIGHT}px !important;
  }}

  /* back-to-top */
  .back-top {{
    position: fixed; bottom: 22px; right: 26px;
    width: 38px; height: 38px;
    background: #262d3d; color: #7aa2f7;
    border: 1px solid #3a4560; border-radius: 50%; cursor: pointer;
    font-size: 1.1rem; line-height: 38px; text-align: center;
    opacity: 0; transition: opacity .2s;
  }}
  .back-top.show {{ opacity: 1; }}
  .back-top:hover {{ background: #3a4560; }}
</style>
</head>
<body>

<div class="header">
  <h1>AI Tool Adoption &amp; Usage Trends</h1>
  <p class="meta">
    <span>Tools: {', '.join(tools)}</span>
    <span>Months: {', '.join(months)}</span>
  </p>
</div>

<nav class="toc">
  <h2>Sections</h2>
  <ul>
{toc}
  </ul>
</nav>

{charts_body}

<button class="back-top" id="backTop" onclick="window.scrollTo({{top:0,behavior:'smooth'}})">&#x2191;</button>
<script>
  window.addEventListener('scroll', function() {{
    document.getElementById('backTop').classList.toggle('show', window.scrollY > 400);
  }});
  window.addEventListener('load', function() {{
    window.dispatchEvent(new Event('resize'));
  }});
</script>
</body>
</html>
"""
    path.write_text(html, encoding="utf-8")


def _card_html(entry: ChartEntry) -> str:
    color = _color_for_tool(entry["tool"])
    return (
        f'    <div class="card">'
        f'      <div class="card-hdr">'
        f'        <span class="tool-badge" style="background:{color}">{entry["tool"]}</span>'
        f'      </div>'
        f'      <div class="card-chart">{entry["html"]}</div>'
        f'    </div>'
    )
