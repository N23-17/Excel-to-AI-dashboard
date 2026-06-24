# Excel to AI Dashboard

Turn Excel or CSV files into an interactive HTML dashboard with Python, Pandas,
and Plotly.

## Features

- Load `.xlsx`, `.xlsm`, and `.csv` files
- Clean empty rows and columns automatically
- Detect date columns for trend charts
- Generate category and numeric summaries
- Export a standalone `dashboard.html` file

## Example Output

![Dashboard preview](Screenshots/Chart_1.png)
![Dashboard preview](Screenshots/Chart_2.png)

## Requirements

- Python 3.8+
- pandas
- plotly
- openpyxl

Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

Run the tool with the included sample workbook:

```bash
python Sheetsense_AI.py sample_data.xlsx
```

Choose a custom output path:

```bash
python Sheetsense_AI.py sample_data.xlsx --output Demo/dashboard.html
```

Choose an Excel sheet by name or index:

```bash
python Sheetsense_AI.py sample_data.xlsx --sheet 0
```

Open the generated HTML file in your browser to view the dashboard.

## Want AI Insights?

The full version includes AI-ready JSON export, a premium prompt pack, and a
faster workflow:

https://imranntenje.gumroad.com/l/sheetsense-ai

## Built By

Imran Ntenje - building simple tools that turn complexity into clarity.
