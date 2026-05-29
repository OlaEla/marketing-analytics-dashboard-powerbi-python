# Marketing Analytics Dashboard

## Overview

End-to-end marketing analytics project that evaluates campaign effectiveness, customer journey, and ROMI (Return on Marketing Investment). The solution integrates data from Google Analytics and CRM, builds conversion touchpoint chains, and provides an interactive Power BI dashboard.

**Business value:** Helps marketing teams identify which channels, campaigns, and user segments drive actual sales, optimize budget allocation, and forecast future conversions.

**Tech stack:** Python (pandas, numpy, statsmodels, seaborn), Excel, Power BI

## Files

- `powerbi/marketing_dashboard.pbix` — interactive dashboard
- `notebooks/analysis.ipynb` — ETL and analysis notebook
- `data/processed_data.csv` — final processed dataset

## Key Features

- **Data integration** — Merges web analytics (GA) and CRM data using Client ID
- **Touchpoint analysis** — Groups user sessions into conversion chains with UTM parameters
- **ROMI calculation** — Custom methodology for marketing ROI by channel
- **Forecasting** — Time series prediction of conversions (Holt's linear trend)
- **Interactive dashboard** — Power BI with 10+ visualizations and filters

## ROMI Methodology

**Note:** True ROMI requires marketing cost data (not available). This analysis uses a simplified metric: **average revenue per sale by channel** as a proxy for channel quality and potential ROI.

## Dashboard Highlights

| Metric | Description |
|--------|-------------|
| Conversion touchpoints | Session chains colored by purchase completion |
| Sales map | Geographic distribution by region and city |
| Sales funnel matrix | By model and brand |
| ROMI by channel | Average revenue per sale per source |
| Daily sales | With target thresholds (customizable) |
| Touchpoint count | Average touches per conversion chain |

## Data Processing Pipeline

```text
Raw data (GA + CRM)
↓
ETL (Python)
├── Brand/model extraction from marketing data
├── Currency rates (Central Bank RF API)
├── Touchpoint chain building
└── ROMI & revenue calculation
↓
Processed dataset
↓
Power BI Dashboard
```

## Results (Sample)

| Metric | Value |
|--------|-------|
| Top converting region | Krasnodar Krai (23,849 requests) |
| Most effective device | Mobile (59.6% of visits) |
| Best performing source (conversion rate) | rambler.ru (83.3%) |
| Highest ROMI source (simplified) | mersedes.ru (5.7M RUB avg sale) |
| Forecast trend (Feb 18–29) | 550 → 306 daily conversions |
| Q1 revenue projection | ~25.7M RUB |

## Quick Start

```bash
# Clone repository
git clone https://github.com/OlaEla/marketing-analytics-dashboard
cd marketing-analytics-dashboard

# Install dependencies
pip install -r requirements.txt

# Run ETL pipeline (if needed)
python scripts/run_etl.py

# Open Power BI dashboard
# File → Open → powerbi/marketing_dashboard.pbix
```

## Author

Oleg Elagin — [GitHub](https://github.com/OlaEla)
