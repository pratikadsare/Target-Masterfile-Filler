import re
from io import BytesIO
import difflib
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from typing import List, Dict, Optional
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Target Masterfile Filler — Interactive Mapping", layout="wide")

# ---- Preserve scroll position across reruns (prevents jumping to top) ----
def inject_scroll_restoration():
    components.html(
        """
        <script>
        const KEY = 'scrollY';
        window.addEventListener('load', function() {
          const y = sessionStorage.getItem(KEY);
          if (y !== null) { window.scrollTo(0, parseFloat(y)); }
        });
        window.addEventListener('beforeunload', function() {
          sessionStorage.setItem(KEY, window.scrollY);
        });
        </script>
        """,
        height=0
    )
inject_scroll_restoration()

# --- Header ---
st.markdown("<h1 style='text-align: center;'>🧩 Target Masterfile Filler — Interactive Mapping</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; font-style: italic;'>Innovating with AI Today ⏐ Leading Automation Tomorrow</h4>", unsafe_allow_html=True)
st.caption("First sheet only, row 1 = headers, row 2 preserved, data written from row 3. Other sheets unchanged. Duplicates in Partner SKU/Barcode highlighted in Excel.")

# ----------------- Helpers -----------------
def _norm_key(s: str) -> str:
    s = str(s or "").strip().lower()
    s = s.replace("&", "and")
    return re.sub(r"[^a-z0-9]+", "", s)

def get_template_headers_from_first_sheet(uploaded_template) -> List[str]:
    wb = load_workbook(
        filename=BytesIO(uploaded_template.getbuffer()),
        data_only=False,
        keep_vba=uploaded_template.name.lower().endswith(".xlsm"),
    )
    ws = wb.worksheets[0]
    headers = []
    max_col = ws.max_column or 1
    for c in range(1, max_col + 1):
        v = ws.cell(row=1, column=c).value
        if v is None:
            continue
        s = str(v).strip()
        if s:
            headers.append(s)
    return headers

def read_raw(uploaded_file, sheet: Optional[str]) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame()
    name = uploaded_file.name.lower()
    if name.endswith((".csv", ".txt")):
        return pd.read_csv(uploaded_file)
    if sheet is not None:
        return pd.read_excel(uploaded_file, sheet_name=sheet, engine="openpyxl")
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    first = xls.sheet_names[0]
    return pd.read_excel(xls, sheet_name=first, engine="openpyxl")

def build_header_index_first_sheet(ws, header_row: int = 1) -> Dict[str, int]:
    headers = {}
    max_col = ws.max_column or 1
    for c in range(1, max_col + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            continue
        key = _norm_key(v)
        if key and key not in headers:
            headers[key] = c
    return headers

def highlight_duplicates(ws, header_map: Dict[str, int], header_labels: List[str], start_row: int = 3):
    """Excel-side highlight: mark duplicate cell values yellow for the given headers (from row 3 down)."""
    dup_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for hdr in header_labels:
        key = _norm_key(hdr)
        col_idx = header_map.get(key)
        if not col_idx:
            continue
        counts = {}
        max_row = ws.max_row or start_row
        for r in range(start_row, max_row + 1):
            v = ws.cell(row=r, column=col_idx).value
            if v is None: 
                continue
            s = str(v).strip()
            if not s:
                continue
            counts[s.upper()] = counts.get(s.upper(), 0) + 1
        if counts:
            for r in range(start_row, max_row + 1):
                v = ws.cell(row=r, column=col_idx).value
                if v is None:
                    continue
                s = str(v).strip()
                if not s:
                    continue
                if counts.get(s.upper(), 0) > 1:
                    ws.cell(row=r, column=col_idx).fill = dup_fill

def fill_first_sheet_by_headers(template_bytes: BytesIO, mapping_df: pd.DataFrame, raw_df: pd.DataFrame, template_filename: str) -> BytesIO:
    """
    Write values ONLY on the first sheet using headers in row 1 (normalized match).
    Start writing from row 3. After filling, highlight duplicates in 'Partner SKU' and 'Barcode'.
    Other sheets remain untouched (names, formatting, formulas, merges).
    """
    keep_vba = template_filename.lower().endswith(".xlsm")
    wb = load_workbook(filename=template_bytes, data_only=False, keep_vba=keep_vba)
    ws = wb.worksheets[0]  # FIRST sheet only

    tpl_header_to_col = build_header_index_first_sheet(ws, header_row=1)
    if not tpl_header_to_col:
        raise ValueError("No headers found in row 1 of the first sheet. Please ensure row 1 contains headers.")

    raw_norm = {c.strip().lower(): c for c in raw_df.columns}

    # Build mapping pairs
    pairs = []
    missing_raw, missing_tpl = [], []
    for _, r in mapping_df.iterrows():
        raw_hdr_lc = str(r["raw_header"]).strip().lower()
        tpl_hdr_norm = _norm_key(r["template_header"])
        if not raw_hdr_lc:
            continue
        raw_col_name = raw_norm.get(raw_hdr_lc)
        col_idx = tpl_header_to_col.get(tpl_hdr_norm)
        if raw_col_name is None:
            missing_raw.append(r["raw_header"]); continue
        if col_idx is None:
            missing_tpl.append(r["template_header"]); continue
        pairs.append((raw_col_name, col_idx))

    if missing_tpl:
        st.warning("Template headers not found in row 1 (skipped): " + ", ".join(sorted(set(missing_tpl))))
    if missing_raw:
        st.warning("RAW columns missing (skipped): " + ", ".join(sorted(set(missing_raw))))

    # Write starting from row 3
    start_row = 3
    for i, (_, raw_row) in enumerate(raw_df.iterrows()):
        out_row = start_row + i
        for raw_col_name, col_idx in pairs:
            val = raw_row[raw_col_name]
            ws.cell(row=out_row, column=col_idx, value=("" if pd.isna(val) else val))

    # Highlight duplicates in specific columns
    highlight_duplicates(ws, tpl_header_to_col, header_labels=["Partner SKU", "Barcode"], start_row=start_row)

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def sanitize_filename(name: str, keep_xlsm: bool) -> str:
    name = (name or "").strip()
    if not name:
        return "filled_masterfile.xlsm" if keep_xlsm else "filled_masterfile.xlsx"
    name = re.sub(r'[\\/*?:"<>|]+', "_", name)
    if keep_vba and not name.lower().endswith(".xlsm"):
        name += ".xlsm"
    elif not keep_vba and not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name

# ----------------- State -----------------
def _init_state():
    defaults = {
        "template_file": None,
        "template_headers": [],
        "template_sig": "",
        "raw_df": pd.DataFrame(),
        "raw_headers": [],
        "raw_sig": "",
        "edit_tick": 0,
        "row_edit_tick": {},  # row index -> last edit tick
        "auto_resolve_on_download": True,  # ON by default
        # mapping values for each template header will be stored as:
        # st.session_state[f"map_{_norm_key(template_header)}"] = <raw header or "">
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()

# ----------------- UI Tabs -----------------
tab1, tab2, tab3 = st.tabs([
    "1) Upload Masterfile Template",
    "2) Upload Raw Data",
    "3) Interactive Mapping & Download"
])

# ---- Tab 1: Template ----
with tab1:
    st.subheader("Upload Masterfile Template (XLSX or XLSM)")
    template_file = st.file_uploader("Masterfile (Excel .xlsx or .xlsm)", type=["xlsx","xlsm"], key="template_file_upl")

    if template_file is not None:
        headers = get_template_headers_from_first_sheet(template_file)
        new_sig = "|".join(headers)
        st.success(f"Template loaded. Headers in row 1 (first sheet): {len(headers)} found.")
        st.dataframe(pd.DataFrame({"Template Headers (row 1)": headers}), use_container_width=True)

        # If template headers actually changed, reset mapping keys for old headers (soft reset)
        if new_sig != st.session_state.template_sig:
            # We don't forcibly clear your choices; we just update the signature and header list.
            st.session_state.template_headers = headers
            st.session_state.template_sig = new_sig

        st.session_state.template_file = template_file

# ---- Tab 2: Raw ----
with tab2:
    st.subheader("Upload Raw Data (CSV/XLSX)")
    raw_file = st.file_uploader("Raw file", type=["csv","xlsx"], key="raw_file_upl")
    raw_sheet = None
    if raw_file is not None and raw_file.name.lower().endswith(".xlsx"):
        try:
            sheets = pd.ExcelFile(raw_file, engine="openpyxl").sheet_names
            if sheets:
                raw_sheet = st.selectbox("Select RAW sheet", options=sheets, index=0, key="raw_sheet_select")
        except Exception as e:
            st.error(f"Could not read RAW Excel: {e}")

    raw_df = read_raw(raw_file, raw_sheet)
    if not raw_df.empty:
        new_sig = "|".join(list(raw_df.columns.astype(str)))
        st.session_state.raw_df = raw_df
        st.session_state.raw_headers = list(raw_df.columns.astype(str))

        # When raw headers change, invalidate any per-row selection that doesn't exist anymore
        if new_sig != st.session_state.raw_sig:
            for th in st.session_state.template_headers:
                key = f"map_{_norm_key(th)}"
                val = st.session_state.get(key, "")
                if val and val not in st.session_state.raw_headers:
                    st.session_state[key] = ""
            st.session_state.raw_sig = new_sig

        st.success(f"RAW loaded. Columns found: {len(st.session_state.raw_headers)}")
        st.dataframe(pd.DataFrame({"Raw Headers": st.session_state.raw_headers}), use_container_width=True)

# ---- Tab 3: Interactive Mapping & Download ----
with tab3:
    st.subheader("Simple Mapping (stable)")
    if not st.session_state.template_headers:
        st.info("Upload a Template in Tab 1 to start.")
    if st.session_state.raw_df.empty:
        st.info("Upload RAW data in Tab 2 to proceed.")

    left, mid, right = st.columns([1,2,1], gap="large")

    with left:
        st.markdown("**Template Headers (row 1)**")
        st.dataframe(pd.DataFrame({"Template": st.session_state.template_headers}), use_container_width=True, height=420)

    with right:
        st.markdown("**Tools**")
        # Auto-map exact
        if st.button("Auto‑map (exact names)"):
            raw_norm_map = {_norm_key(h): h for h in st.session_state.raw_headers}
            for th in st.session_state.template_headers:
                cand = raw_norm_map.get(_norm_key(th), "")
                if cand and not st.session_state.get(f"map_{_norm_key(th)}"):
                    st.session_state[f"map_{_norm_key(th)}"] = cand
            st.toast("Exact auto‑map applied.", icon="✅")
        # Auto-map fuzzy
        fuzz_thresh = st.slider("Fuzzy match threshold", 0, 100, 80, 1, help="Percent similarity; higher = stricter")
        if st.button("Auto‑map (fuzzy)"):
            for th in st.session_state.template_headers:
                key = f"map_{_norm_key(th)}"
                if st.session_state.get(key, ""):
                    continue
                tpl_norm = _norm_key(th)
                cands = [(h, difflib.SequenceMatcher(None, tpl_norm, _norm_key(h)).ratio()) for h in st.session_state.raw_headers]
                cands.sort(key=lambda x: x[1], reverse=True)
                if cands and int(cands[0][1]*100) >= fuzz_thresh:
                    st.session_state[key] = cands[0][0]
            st.toast("Fuzzy auto‑map applied.", icon="✅")

        st.markdown("---")
        # Import mapping
        imp = st.file_uploader("Import mapping (CSV/XLSX)", type=["csv","xlsx"], key="import_map_upl")
        if imp is not None:
            try:
                if imp.name.lower().endswith(".csv"):
                    imp_df = pd.read_csv(imp)
                else:
                    imp_df = pd.read_excel(imp, engine="openpyxl")
                cols_norm = {_norm_key(c): c for c in imp_df.columns}
                tpl_col = next((cols_norm[c] for c in ["template","templateheader","headerofmasterfiletemplate","target","to"] if c in cols_norm), None)
                raw_col = next((cols_norm[c] for c in ["raw","rawheader","headerofrowsheet","source","from"] if c in cols_norm), None)
                if not tpl_col or not raw_col:
                    st.warning("Import needs 'Template' and 'Raw' columns.")
                else:
                    for _, row in imp_df[[tpl_col, raw_col]].dropna().iterrows():
                        th = str(row[tpl_col])
                        rv = str(row[raw_col])
                        if th in st.session_state.template_headers and rv in st.session_state.raw_headers:
                            st.session_state[f"map_{_norm_key(th)}"] = rv
                    st.success("Mapping imported.")
            except Exception as e:
                st.error(f"Failed to import mapping: {e}")

        # Export mapping
        if st.button("Export mapping (CSV)"):
            rows = []
            for th in st.session_state.template_headers:
                rows.append({"Template": th, "Raw": st.session_state.get(f"map_{_norm_key(th)}", "")})
            out = pd.DataFrame(rows).to_csv(index=False).encode("utf-8")
            st.download_button("Download mapping.csv", data=out, file_name="mapping.csv", mime="text/csv")

        st.markdown("---")
        st.session_state.auto_resolve_on_download = st.checkbox(
            "Auto‑resolve duplicate Raw selections on download (latest edit wins)",
            value=st.session_state.auto_resolve_on_download
        )

    with mid:
        st.markdown("**Pick a Raw column for each Template header**")
        options = [""] + st.session_state.raw_headers
        # Render a compact two-column selector list
        for th in st.session_state.template_headers:
            key = f"map_{_norm_key(th)}"
            current = st.session_state.get(key, "")
            if current not in st.session_state.raw_headers:
                current = ""
            st.selectbox(
                label=th,
                options=options,
                index=(options.index(current) if current in options else 0),
                key=key,
            )

        st.markdown("---")
        default_name = "filled_masterfile.xlsm" if (st.session_state.template_file and st.session_state.template_file.name.lower().endswith(".xlsm")) else "filled_masterfile.xlsx"
        output_name = st.text_input("Output file name", value=default_name, help="Enter the name to use for the downloaded file (extension will be enforced).")

        can_process = (
            st.session_state.template_file is not None and
            not st.session_state.raw_df.empty and
            len(st.session_state.template_headers) > 0
        )
        if not can_process:
            st.info("Upload Template (Tab 1) and Raw (Tab 2) before processing.")

        if st.button("⚙️ Process & Download", type="primary", disabled=not can_process):
            try:
                # Build mapping_df from the per-row selectboxes
                rows = []
                for th in st.session_state.template_headers:
                    raw_sel = st.session_state.get(f"map_{_norm_key(th)}", "")
                    if raw_sel:
                        rows.append({"raw_header": raw_sel, "template_header": th})
                mapping_df = pd.DataFrame(rows)
                if mapping_df.empty:
                    st.error("No mappings selected.")
                    st.stop()

                # Resolve duplicates (same Raw chosen for multiple Templates) at download time
                if st.session_state.auto_resolve_on_download:
                    # keep the row with the highest most recent selection (we don't track per-row ticks here,
                    # so we keep the last occurrence in the list which corresponds to the last UI control rendered)
                    seen = set()
                    dedup_rows = []
                    for r in reversed(rows):  # reverse to keep last selection
                        k = r["raw_header"]
                        if k in seen:  # already kept a later one
                            continue
                        seen.add(k)
                        dedup_rows.append(r)
                    dedup_rows.reverse()
                    mapping_df = pd.DataFrame(dedup_rows)

                # Fill workbook
                out_bytes = fill_first_sheet_by_headers(
                    template_bytes=BytesIO(st.session_state.template_file.getbuffer()),
                    mapping_df=mapping_df,
                    raw_df=st.session_state.raw_df,
                    template_filename=st.session_state.template_file.name
                )

                keep_vba = st.session_state.template_file.name.lower().endswith(".xlsm")
                # enforce extension
                name = (output_name or "").strip()
                if not name:
                    name = "filled_masterfile.xlsm" if keep_vba else "filled_masterfile.xlsx"
                if keep_vba and not name.lower().endswith(".xlsm"):
                    name += ".xlsm"
                elif not keep_vba and not name.lower().endswith(".xlsx"):
                    name += ".xlsx"

                st.success("Done! Your updated masterfile is ready.")
                st.download_button(
                    label="⬇️ Download Updated Masterfile",
                    data=out_bytes.getvalue(),
                    file_name=name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.exception(e)
