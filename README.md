# Masterfile Filler — First Sheet, Row 3 Writer (Streamlit)

This app updates **only the FIRST sheet** of your Excel masterfile, using **Row 1** as headers, preserves **Row 2**, and fills data **starting at Row 3**. **Other sheets** remain intact: names, styles, colors, borders, merged cells, and formulas.

## What it does
- Maps RAW columns → Template columns using a simple **2‑column mapping**:
  - **header of row sheet** (RAW)
  - **header of masterfile template** (must exist on Row 1 of the first sheet)
- Writes values from RAW rows to the template starting at **Row 3**.
- Keeps other sheets untouched.

## Use it (Streamlit Cloud)
1. Push this repo to GitHub.
2. Deploy on Streamlit Community Cloud → **New app** → main file: `streamlit_app.py`.
3. Share the URL.

## Local run
```
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Files included
- `streamlit_app.py` — browser UI
- `mapping_template.xlsx` — 2‑column mapping
- `sample_data/` — sample 3‑sheet template, raw, and mapping

MIT License © 2025
