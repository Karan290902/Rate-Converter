# AVIVA GCL HL → LAP Converter

## Features

- Upload AVIVA GCL HL Excel
- Automatic Wellness Calculation
- Automatic Net Calculation
- Automatic Gross Calculation
- Download Final Excel

## Formula Logic

Wellness = Premium × 5%

Net = Premium + Wellness

Gross = Net × 1.18

## Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
