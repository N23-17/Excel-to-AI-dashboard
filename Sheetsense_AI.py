# SheetSense AI - Excel to AI Dashboard
# Author: Imran Ntenje

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
from pandas.api.types import is_datetime64_any_dtype, is_numeric_dtype


def load_file(file_path: str | Path, sheet_name: str | int | None = None) -> pd.DataFrame:
    """Load an Excel or CSV file into a DataFrame."""
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path)

    if suffix in {".xlsx", ".xlsm"}:
        return pd.read_excel(
            path,
            sheet_name=0 if sheet_name is None else sheet_name,
            engine="openpyxl",
        )

    raise ValueError("Unsupported file type. Please use .xlsx, .xlsm, or .csv.")


def auto_clean(df: pd.DataFrame) -> pd.DataFrame:
    """Clean common spreadsheet issues without changing the original DataFrame."""
    cleaned = df.copy()

    cleaned = cleaned.dropna(how="all").dropna(axis=1, how="all")
    cleaned.columns = [str(column).strip() for column in cleaned.columns]

    for column in cleaned.columns:
        if cleaned[column].dtype == "object":
            converted_dates = pd.to_datetime(cleaned[column], errors="coerce")
            if converted_dates.notna().mean() > 0.5:
                cleaned[column] = converted_dates
                continue

        if is_numeric_dtype(cleaned[column]):
            cleaned[column] = cleaned[column].fillna(0)

    return cleaned


def _choose_date_column(df: pd.DataFrame) -> str | None:
    for column in df.columns:
        if is_datetime64_any_dtype(df[column]):
            return column
    return None


def _choose_numeric_columns(df: pd.DataFrame) -> list[str]:
    return [column for column in df.columns if is_numeric_dtype(df[column])]


def _choose_category_columns(df: pd.DataFrame) -> list[str]:
    return [
        column
        for column in df.columns
        if not is_numeric_dtype(df[column]) and not is_datetime64_any_dtype(df[column])
    ]


def build_figures(df: pd.DataFrame) -> list:
    """Build a useful default set of charts from the available columns."""
    figures = []
    numeric_columns = _choose_numeric_columns(df)
    category_columns = _choose_category_columns(df)
    date_column = _choose_date_column(df)

    if numeric_columns:
        value_column = numeric_columns[-1]

        if date_column:
            trend_df = df.sort_values(date_column)
            figures.append(
                px.line(
                    trend_df,
                    x=date_column,
                    y=value_column,
                    markers=True,
                    title=f"{value_column} over time",
                )
            )
        else:
            figures.append(
                px.bar(
                    df.head(20).reset_index(names="Record"),
                    x="Record",
                    y=value_column,
                    title=f"First 20 records by {value_column}",
                )
            )

    if category_columns and numeric_columns:
        category_column = category_columns[0]
        value_column = numeric_columns[-1]
        grouped = (
            df.groupby(category_column, dropna=False)[value_column]
            .sum()
            .sort_values(ascending=False)
            .head(10)
            .reset_index()
        )
        figures.append(
            px.bar(
                grouped,
                x=category_column,
                y=value_column,
                title=f"Top {category_column} by {value_column}",
            )
        )

    if category_columns:
        category_column = category_columns[0]
        counts = df[category_column].fillna("Unknown").value_counts().head(10)
        figures.append(
            px.pie(
                values=counts.values,
                names=counts.index,
                title=f"{category_column} distribution",
            )
        )

    return figures


def generate_dashboard(df: pd.DataFrame, output_html: str | Path = "dashboard.html") -> Path:
    """Generate a standalone interactive HTML dashboard."""
    output_path = Path(output_html)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    figures = build_figures(df)

    chart_html = []
    for index, figure in enumerate(figures):
        figure.update_layout(
            template="plotly_white",
            margin=dict(l=40, r=30, t=70, b=45),
            height=430,
        )
        chart_html.append(
            f'<section class="chart"><h2>Chart {index + 1}</h2>'
            f'{figure.to_html(full_html=False, include_plotlyjs=index == 0)}</section>'
        )

    if not chart_html:
        chart_html.append(
            '<section class="chart"><h2>No charts generated</h2>'
            "<p>Add at least one numeric or category column to generate visuals.</p></section>"
        )

    preview_rows = df.head(10).to_html(index=False, classes="preview-table", border=0)
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    html = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>SheetSense AI Dashboard</title>
  <style>
    :root {{
      color-scheme: light;
      --bg: #f6f8fb;
      --panel: #ffffff;
      --ink: #162033;
      --muted: #667085;
      --line: #d9e2ec;
      --accent: #2563eb;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      background: var(--bg);
      color: var(--ink);
      font-family: Arial, Helvetica, sans-serif;
      line-height: 1.5;
    }}
    header, main {{
      width: min(1120px, calc(100% - 32px));
      margin: 0 auto;
    }}
    header {{
      padding: 40px 0 24px;
    }}
    h1, h2 {{
      margin: 0;
      letter-spacing: 0;
    }}
    h1 {{
      font-size: clamp(2rem, 4vw, 3.5rem);
    }}
    h2 {{
      font-size: 1.1rem;
      margin-bottom: 16px;
    }}
    .meta {{
      color: var(--muted);
      margin: 8px 0 0;
    }}
    .stats {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 12px;
      margin-bottom: 18px;
    }}
    .stat, .chart, .preview {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: 0 12px 30px rgba(22, 32, 51, 0.06);
    }}
    .stat {{
      padding: 16px;
    }}
    .stat span {{
      color: var(--muted);
      display: block;
      font-size: 0.85rem;
    }}
    .stat strong {{
      display: block;
      font-size: 1.55rem;
      margin-top: 4px;
    }}
    .chart, .preview {{
      padding: 18px;
      margin-bottom: 18px;
      overflow: hidden;
    }}
    .table-wrap {{
      overflow-x: auto;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      min-width: 720px;
    }}
    th, td {{
      border-bottom: 1px solid var(--line);
      padding: 10px 12px;
      text-align: left;
      white-space: nowrap;
    }}
    th {{
      color: var(--muted);
      font-size: 0.82rem;
      text-transform: uppercase;
    }}
    footer {{
      color: var(--muted);
      padding: 18px 0 36px;
      text-align: center;
    }}
  </style>
</head>
<body>
  <header>
    <h1>SheetSense AI Dashboard</h1>
    <p class="meta">Generated on {generated_at}</p>
  </header>
  <main>
    <section class="stats" aria-label="Dataset summary">
      <div class="stat"><span>Rows</span><strong>{len(df):,}</strong></div>
      <div class="stat"><span>Columns</span><strong>{len(df.columns):,}</strong></div>
      <div class="stat"><span>Charts</span><strong>{len(figures):,}</strong></div>
    </section>
    {''.join(chart_html)}
    <section class="preview">
      <h2>Data preview</h2>
      <div class="table-wrap">{preview_rows}</div>
    </section>
  </main>
  <footer>Built with Python, Pandas, and Plotly.</footer>
</body>
</html>
"""

    output_path.write_text(html, encoding="utf-8")
    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Turn an Excel or CSV file into an interactive HTML dashboard."
    )
    parser.add_argument("file", nargs="?", help="Path to the Excel or CSV file.")
    parser.add_argument(
        "-o",
        "--output",
        default="dashboard.html",
        help="Dashboard output path. Defaults to dashboard.html.",
    )
    parser.add_argument(
        "-s",
        "--sheet",
        default=None,
        help="Excel sheet name or index. Defaults to the first sheet.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    file_path = args.file or input("Enter Excel/CSV filename (for example, sales.xlsx): ").strip()

    sheet_name = args.sheet
    if isinstance(sheet_name, str) and sheet_name.isdigit():
        sheet_name = int(sheet_name)

    try:
        df = load_file(file_path, sheet_name=sheet_name)
        cleaned = auto_clean(df)
        dashboard_path = generate_dashboard(cleaned, args.output)
    except Exception as error:
        raise SystemExit(f"Error: {error}") from error

    print(f"Loaded {len(cleaned):,} rows and {len(cleaned.columns):,} columns.")
    print(f"Dashboard saved to {dashboard_path}")


if __name__ == "__main__":
    main()
