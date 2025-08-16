
# Vans Data Interactive Dashboard (Enhanced + PDF Export)

## Features
- Password-protected access (set STREAMLIT_DASH_PASSWORD)
- Global filters (Company, Age, Employment, Area)
- KPI strip (Deliveries, Income, Insurance, etc.)
- PivotTable.js for slicing
- Preset pivots
- Interactive Plotly charts
- 📑 Export to PDF (KPIs + charts)

## Run locally
```bash
pip install -r requirements.txt
export STREAMLIT_DASH_PASSWORD="your-password"   # optional
streamlit run app.py
```

## Free Hosting
- **Streamlit Cloud** or **Hugging Face Spaces**
- Add env var `STREAMLIT_DASH_PASSWORD` for secure access
