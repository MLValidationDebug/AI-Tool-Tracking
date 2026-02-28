# AI Tool Adoption Dashboard

Interactive dashboard for tracking month-over-month AI tool adoption and active usage across teams. Parses an Excel workbook, extracts function-level metrics, and generates a self-contained Plotly HTML dashboard with dark-themed charts.

## Features

- **Automatic Excel parsing** — detects month blocks, header rows, and function-level data from each tool sheet
- **Interactive Plotly charts** — Adoption Rate & Active Rate trend lines per tool per function; dual-axis bar charts for Microsoft Copilot actions
- **All-in-one dashboard** — single `dashboard.html` with table of contents, responsive grid layout, and a dark UI theme
- **Flask web server** — serve the dashboard locally with a one-click refresh endpoint to re-parse and regenerate
- **Extensible** — add a new tool by creating a sheet in the workbook and registering it in `parser.py → TOOL_SHEETS`

## Tracked Tools

| Sheet name | Parser | Metrics |
|---|---|---|
| **AI Code** (GitHub Copilot, Cline, Cursor, CodeGen, Genie) | `standard` | Adoption Rate, Active Rate |
| **NABU** | `standard` | Adoption Rate, Active Rate |
| **Microsoft Copilot** | `copilot` | Adoption Rate, Active Rate, Total Actions, Avg Action/User |

## Project Structure

```
├── main.py          # CLI entry point — parse workbook & generate charts
├── app.py           # Flask web server — serve & refresh dashboard
├── parser.py        # Excel workbook parser (standard + Copilot layouts)
├── charts.py        # Plotly chart builders & HTML dashboard renderer
├── requirements.txt # Python dependencies
├── Data/            # Place your .xlsx workbook here
└── output/          # Generated at runtime (dashboard + intermediate CSV)
```

`output/` and Python cache folders are generated locally and ignored by git.

## Setup

```bash
pip install -r requirements.txt
```

**Requirements:** Python 3.10+, pandas, openpyxl, plotly, flask

## Usage

### CLI — generate once

```bash
python main.py                              # auto-detects .xlsx in Data/
python main.py -i path/to/file.xlsx -o output
```

### Web server — serve & refresh

```bash
python app.py
```

- Dashboard: `http://localhost:5000/`
- Refresh (re-parse & regenerate): `http://localhost:5000/refresh`

## Output

| File | Description |
|---|---|
| `output/dashboard.html` | All-in-one interactive dashboard |
| `output/extracted_data.csv` | Intermediate tidy CSV for further analysis |

## Data Layout Expected

Each tool sheet should contain month blocks stacked vertically:

```
Feb. 26
  (header row)
  Function | Leader | Total | Adoption | Active | Adoption Rate | Active Rate
  ...data rows...
  Total row

Jan. 26
  (header row)
  ...
```

The parser auto-detects month labels and header rows.
