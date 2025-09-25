# PJM Day-Ahead Prices — Unified HTML Dashboard

Generates a single, aesthetic HTML dashboard + downloadable CSV summaries from an Excel file of PJM zone day-ahead prices.

## Input Format

Excel sheet with:
- Column `Date` (timestamps)
- One column per zone (e.g., `AECO ZONE`, `PJM-RTO ZONE`, …)

Default sheet name: `ZoneWisePrices`.

## Quick Start

```bash
python -m venv .venv
# On macOS/Linux:
source .venv/bin/activate
# On Windows:
# .venv\Scripts\activate

pip install -r requirements.txt

python generate_pjm_dashboard.py \
  --input "PJM_DayAhead_Prices_ZoneWise (1).xlsx" \
  --sheet ZoneWisePrices \
  --output PJM_Unified_Dashboard.html \
  --assets-dir assets \
  --include-plotlyjs cdn
