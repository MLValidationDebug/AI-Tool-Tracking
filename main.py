"""
main.py — Entry point for AI Tool Adoption chart generation.

Usage:
    python main.py                          # uses defaults
    python main.py --input data/MyFile.xlsx --output output
"""

from __future__ import annotations

import argparse
import glob
from pathlib import Path

from parser import parse_workbook
from charts import generate_charts


def find_excel(data_dir: str = "Data") -> Path:
    """Auto-detect the first .xlsx file in the data directory."""
    pattern = str(Path(data_dir) / "*.xlsx")
    files = glob.glob(pattern)
    if not files:
        raise FileNotFoundError(
            f"No .xlsx files found in '{data_dir}/'. "
            "Place your AI Tool Adoption workbook there."
        )
    return Path(files[0])


def main() -> None:
    ap = argparse.ArgumentParser(description="AI Tool Adoption trend charts")
    ap.add_argument(
        "--input", "-i",
        default=None,
        help="Path to the Excel workbook (default: auto-detect in Data/)",
    )
    ap.add_argument(
        "--output", "-o",
        default="output",
        help="Directory for generated charts (default: output/)",
    )
    args = ap.parse_args()

    input_path = Path(args.input) if args.input else find_excel()
    print(f"Reading workbook: {input_path}")

    df = parse_workbook(input_path)
    if df.empty:
        print("No data extracted — check the workbook structure.")
        return

    print(f"\nExtracted {len(df)} records across "
          f"{df['Tool'].nunique()} tools, "
          f"{df['Month'].nunique()} months, "
          f"{df['Function'].nunique()} functions")

    # Save intermediate CSV for inspection
    csv_path = Path(args.output) / "extracted_data.csv"
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(csv_path, index=False)
    print(f"  Intermediate CSV → {csv_path}")

    dashboard = generate_charts(df, args.output)
    print(f"\n✓ Open the dashboard: {dashboard}")


if __name__ == "__main__":
    main()
