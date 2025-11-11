# Target Masterfile Filler — Interactive Mapping

This app adds an **interactive mapping grid** to map your onboarding/raw sheet headers to your masterfile template headers:

- **Left**: Template headers (row 1 of the **first sheet**)  
- **Center**: Mapping grid with **dropdowns** per template row (choose a Raw column)  
- **Right**: Raw/onboarding headers  

## Key behaviors
- First sheet only is edited; **row 1 = headers**, **row 2 preserved**, writes from **row 3**.
- All other sheets remain **unchanged** (names, formats, merges, formulas).
- After filling, **duplicates** in **Partner SKU** and **Barcode** are **highlighted** (yellow) within their columns.
- Filename input lets you **name the output** before download.
- Tools: **auto-map (exact)**, **auto-map (fuzzy)**, **uniqueness toggle**, **import/export mapping**.

## Deploy on Streamlit Cloud
1. Push this repo to GitHub.
2. New app → main file: `streamlit_app.py`.
3. Deploy and share the URL.

## Local run
```
pip install -r requirements.txt
streamlit run streamlit_app.py
```

MIT License © 2025
