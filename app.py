"""
app.py — Flask web service for the AI Tool Adoption Dashboard.

Endpoints:
  GET  /            → serves the dashboard
  POST /refresh     → re-parses the workbook and regenerates the dashboard
"""

from __future__ import annotations

import glob
from pathlib import Path

from flask import Flask, send_file, redirect, url_for, flash, render_template_string

from parser import parse_workbook
from charts import generate_charts

app = Flask(__name__)
app.secret_key = "ai-usage-dashboard"

DATA_DIR = "Data"
OUTPUT_DIR = "output"
DASHBOARD = Path(OUTPUT_DIR) / "dashboard.html"


def _find_excel() -> Path:
    files = glob.glob(str(Path(DATA_DIR) / "*.xlsx"))
    if not files:
        raise FileNotFoundError(f"No .xlsx files found in '{DATA_DIR}/'")
    return Path(files[0])


def _rebuild() -> str | None:
    """Regenerate the dashboard. Returns None on success or an error message."""
    try:
        wb = _find_excel()
        df = parse_workbook(wb)
        if df.empty:
            return "No data extracted — check the workbook structure."
        generate_charts(df, OUTPUT_DIR)
        # Also save intermediate CSV
        csv_path = Path(OUTPUT_DIR) / "extracted_data.csv"
        df.to_csv(csv_path, index=False)
        return None
    except Exception as e:
        return str(e)


# --- Routes ----------------------------------------------------------------

@app.route("/")
def index():
    """Serve the dashboard HTML."""
    if not DASHBOARD.exists():
        # Auto-build on first visit
        err = _rebuild()
        if err:
            return render_template_string(
                "<h2>Dashboard not available</h2><p>{{ err }}</p>"
                '<p><a href="/refresh">Try refreshing</a></p>',
                err=err,
            ), 500

    return send_file(DASHBOARD.resolve(), mimetype="text/html")


@app.route("/refresh", methods=["GET", "POST"])
def refresh():
    """Re-parse the workbook and regenerate the dashboard."""
    err = _rebuild()
    if err:
        return render_template_string(
            "<h2>Refresh failed</h2><p>{{ err }}</p>"
            '<p><a href="/">Back to dashboard</a></p>',
            err=err,
        ), 500
    return redirect(url_for("index"))


# --- Entry point -----------------------------------------------------------

if __name__ == "__main__":
    # Build once at startup so the dashboard is ready
    print("Building dashboard …")
    startup_err = _rebuild()
    if startup_err:
        print(f"⚠  {startup_err}")
    else:
        print("✓ Dashboard ready")

    print("\nStarting server at http://localhost:5000")
    print("  → Dashboard : http://localhost:5000/")
    print("  → Refresh   : http://localhost:5000/refresh")
    app.run(host="0.0.0.0", port=5000, debug=True)
