
import io
import re
from io import BytesIO
import difflib
import pandas as pd
import streamlit as st
from typing import List, Dict, Optional, Tuple
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Target Masterfile Filler — Interactive Mapping", layout="wide")

# --- Header ---
st.markdown("<h1 style='text-align: center;'>🧩 Target Masterfile Filler — Interactive Mapping</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; font-style: italic;'>Innovating with AI Today ⏐ Leading Automation Tomorrow</h4>", unsafe_allow_html=True)
st.caption("First sheet only, row 1 = headers, row 2 preserved, data written from row 3. Other sheets unchanged. Duplicates in Partner SKU/Barcode highlighted.")

# ----------------- Helpers -----------------
def _norm_key(s: str) -> str:
    s = str(s or "").strip().lower()
    s = s.replace("&", "and")
    return re.sub(r"[^a-z0-9]+", "", s)

def get_template_headers_from_first_sheet(uploaded_template) -> List[str]:
    """Read row-1 headers from the first sheet using openpyxl (robust to formatting)."""
    try:
        wb = load_workbook(filename=BytesIO(uploaded_template.getbuffer()), data_only=False, keep_vba=uploaded_template.name.lower().endswith(".xlsm"))
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
    except Exception as e:
        st.error(f"Failed to read template headers: {e}")
        return []

def read_raw(uploaded_file, sheet: Optional[str]) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame()
    name = uploaded_file.name.lower()
    try:
        if name.endswith((".csv", ".txt")):
            return pd.read_csv(uploaded_file)
        if sheet is not None:
            return pd.read_excel(uploaded_file, sheet_name=sheet, engine="openpyxl")
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
        first = xls.sheet_names[0]
        return pd.read_excel(xls, sheet_name=first, engine="openpyxl")
    except Exception as e:
        st.error(f"Failed to read RAW: {e}")
        return pd.DataFrame()

def build_header_index_first_sheet(ws, header_row: int = 1) -> Dict[str, int]:
    """Return normalized header -> column index for FIRST sheet (row 1)."""
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
    """Highlight duplicate cells (yellow) for given headers from start_row down."""
    dup_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for hdr in header_labels:
        key = _norm_key(hdr)
        col_idx = header_map.get(key)
        if not col_idx:
            continue
        counts = {}
        max_row = ws.max_row or start_row
        # count
        for r in range(start_row, max_row + 1):
            v = ws.cell(row=r, column=col_idx).value
            if v is None: 
                continue
            s = str(v).strip()
            if not s:
                continue
            counts[s.upper()] = counts.get(s.upper(), 0) + 1
        # highlight
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

    # Header index
    tpl_header_to_col = build_header_index_first_sheet(ws, header_row=1)
    if not tpl_header_to_col:
        raise ValueError("No headers found in row 1 of the first sheet. Please ensure row 1 contains headers.")

    # Raw lookup
    raw_norm = {c.strip().lower(): c for c in raw_df.columns}

    # mapping_df columns: ['raw_header','template_header']
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

    # Write rows from row 3
    start_row = 3
    for i, (_, raw_row) in enumerate(raw_df.iterrows()):
        out_row = start_row + i
        for raw_col_name, col_idx in pairs:
            val = raw_row[raw_col_name]
            ws.cell(row=out_row, column=col_idx, value=("" if pd.isna(val) else val))

    # Highlight duplicates in specific columns
    highlight_duplicates(ws, tpl_header_to_col, header_labels=["Partner SKU", "Barcode"], start_row=start_row)

    # Save
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def sanitize_filename(name: str, keep_xlsm: bool) -> str:
    name = (name or "").strip()
    if not name:
        return "filled_masterfile.xlsm" if keep_xlsm else "filled_masterfile.xlsx"
    # remove illegal filename chars
    name = re.sub(r'[\\/*?:"<>|]+', "_", name)
    # ensure extension
    if keep_xlsm and not name.lower().endswith(".xlsm"):
        name += ".xlsm"
    elif not keep_xlsm and not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name

# ----------------- Session keys -----------------
if "template_file" not in st.session_state:
    st.session_state.template_file = None
if "template_headers" not in st.session_state:
    st.session_state.template_headers = []
if "raw_df" not in st.session_state:
    st.session_state.raw_df = pd.DataFrame()
if "raw_headers" not in st.session_state:
    st.session_state.raw_headers = []
if "mapping_table" not in st.session_state:
    st.session_state.mapping_table = pd.DataFrame(columns=["Template","Raw"])
if "enforce_unique" not in st.session_state:
    st.session_state.enforce_unique = True

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
        st.session_state.template_file = template_file
        template_headers = get_template_headers_from_first_sheet(template_file)
        st.session_state.template_headers = template_headers

        st.success(f"Template loaded. First sheet headers: {len(template_headers)} found.")
        st.dataframe(pd.DataFrame({"Template Headers (row 1)": template_headers}), use_container_width=True)

        # Reinitialize mapping table, preserving matches where possible
        if not template_headers:
            st.session_state.mapping_table = pd.DataFrame(columns=["Template","Raw"])
        else:
            old_map = st.session_state.mapping_table
            prev_map = {}
            if not old_map.empty and "Template" in old_map.columns and "Raw" in old_map.columns:
                for _, r in old_map.iterrows():
                    prev_map[str(r["Template"])] = str(r["Raw"]) if pd.notna(r["Raw"]) else ""
            rows = []
            for th in template_headers:
                rows.append({"Template": th, "Raw": prev_map.get(th, "")})
            st.session_state.mapping_table = pd.DataFrame(rows, columns=["Template","Raw"])

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
        st.session_state.raw_df = raw_df
        st.session_state.raw_headers = list(raw_df.columns.astype(str))
        st.success(f"RAW loaded. Columns: {len(st.session_state.raw_headers)} found.")
        st.dataframe(pd.DataFrame({"Raw Headers": st.session_state.raw_headers}), use_container_width=True)

        # Merge existing mapping with new raw headers (drop non-existing raw choices)
        if not st.session_state.mapping_table.empty:
            mt = st.session_state.mapping_table.copy()
            mt["Raw"] = mt["Raw"].apply(lambda x: x if str(x) in st.session_state.raw_headers else "")
            st.session_state.mapping_table = mt

# ---- Tab 3: Interactive Mapping & Download ----
with tab3:
    st.subheader("Interactive Mapping")
    if not st.session_state.template_headers:
        st.info("Upload a Template in Tab 1 to start.")
    if st.session_state.raw_df.empty:
        st.info("Upload RAW data in Tab 2 to proceed.")

    col_left, col_mid, col_right = st.columns([1, 2, 1], gap="large")

    with col_left:
        st.markdown("**Template Headers (row 1)**")
        st.dataframe(pd.DataFrame({"Template": st.session_state.template_headers}), use_container_width=True, height=420)

    with col_right:
        st.markdown("**Raw Headers**")
        st.dataframe(pd.DataFrame({"Raw": st.session_state.raw_headers}), use_container_width=True, height=420)

        st.markdown("---")
        st.markdown("**Tools**")
        enforce_unique = st.checkbox("Enforce unique Raw selection", value=st.session_state.enforce_unique, help="Prevents the same Raw column from being mapped to multiple Template rows.")
        st.session_state.enforce_unique = enforce_unique

        # Auto-map exact
        if st.button("Auto‑map (exact names)"):
            mt = st.session_state.mapping_table.copy()
            raw_headers = st.session_state.raw_headers
            used = set(mt["Raw"].dropna().astype(str)) if enforce_unique else set()
            raw_norm_map = {_norm_key(h): h for h in raw_headers}
            for i, row in mt.iterrows():
                if str(row["Raw"]):
                    continue
                tpl = str(row["Template"])
                candidate = raw_norm_map.get(_norm_key(tpl))
                if candidate and ((not enforce_unique) or (candidate not in used)):
                    mt.at[i, "Raw"] = candidate
                    used.add(candidate)
            st.session_state.mapping_table = mt

        # Auto-map fuzzy
        fuzz_thresh = st.slider("Fuzzy match threshold", min_value=0, max_value=100, value=80, step=1, help="Percent similarity; higher = stricter")
        if st.button("Auto‑map (fuzzy)"):
            mt = st.session_state.mapping_table.copy()
            raw_headers = st.session_state.raw_headers
            used = set(mt["Raw"].dropna().astype(str)) if enforce_unique else set()
            for i, row in mt.iterrows():
                if str(row["Raw"]):
                    continue
                tpl = str(row["Template"])
                # Find best match
                # Use normalized strings for scoring
                tpl_norm = _norm_key(tpl)
                candidates = [(h, difflib.SequenceMatcher(None, tpl_norm, _norm_key(h)).ratio()) for h in raw_headers]
                candidates.sort(key=lambda x: x[1], reverse=True)
                if candidates and candidates[0][1] * 100 >= fuzz_thresh:
                    best = candidates[0][0]
                    if (not enforce_unique) or (best not in used):
                        mt.at[i, "Raw"] = best
                        used.add(best)
            st.session_state.mapping_table = mt

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
                tpl_col = None
                raw_col = None
                for cand in ["template","templateheader","headerofmasterfiletemplate","target","to"]:
                    if cand in cols_norm: tpl_col = cols_norm[cand]; break
                for cand in ["raw","rawheader","headerofrowsheet","source","from"]:
                    if cand in cols_norm: raw_col = cols_norm[cand]; break
                if not tpl_col or not raw_col:
                    st.warning("Could not find 'Template' and 'Raw' columns in the import file.")
                else:
                    mt = st.DataFrame({"Template": st.session_state.template_headers})
                    # Map imported rows onto our template list
                    imp_small = imp_df[[tpl_col, raw_col]].dropna()
                    imp_small.columns = ["Template","Raw"]
                    # Keep only raw headers that exist
                    imp_small["Raw"] = imp_small["Raw"].astype(str).apply(lambda x: x if x in st.session_state.raw_headers else "")
                    # Merge
                    merged = pd.merge(mt, imp_small, on="Template", how="left")
                    merged["Raw"] = merged["Raw"].fillna("")
                    st.session_state.mapping_table = merged
                    st.success("Mapping imported into the grid.")
            except Exception as e:
                st.error(f"Failed to import mapping: {e}")

        # Export mapping
        if st.button("Export mapping (CSV)"):
            mt = st.session_state.mapping_table.copy()
            out = mt.to_csv(index=False).encode("utf-8")
            st.download_button("Download mapping.csv", data=out, file_name="mapping.csv", mime="text/csv")

    with col_mid:
        st.markdown("**Center Mapping: pick a Raw column for each Template row**")
        if st.session_state.mapping_table.empty:
            st.info("Mapping grid will appear once both Template and Raw are uploaded.")
        else:
            # Build dropdown options
            options = [""] + st.session_state.raw_headers
            # Show data editor
            cfg = {
                "Template": st.column_config.Column(disabled=True),
                "Raw": st.column_config.SelectboxColumn(
                    "Raw",
                    options=options,
                    help="Select the raw column that maps to this template header.",
                )
            }
            edited = st.data_editor(
                st.session_state.mapping_table,
                column_config=cfg,
                hide_index=True,
                use_container_width=True,
                num_rows="fixed",
                key="mapping_editor"
            )
            # Enforce uniqueness if selected
            if st.session_state.enforce_unique:
                # If a raw value appears multiple times (excluding empty), keep first occurrence, blank the rest
                used = set()
                for i in range(len(edited)):
                    val = str(edited.loc[i, "Raw"]) if pd.notna(edited.loc[i, "Raw"]) else ""
                    if not val:
                        continue
                    if val in used:
                        edited.loc[i, "Raw"] = ""
                    else:
                        used.add(val)
            st.session_state.mapping_table = edited

        st.markdown("---")
        # File name input
        default_name = "filled_masterfile.xlsm" if (st.session_state.template_file and st.session_state.template_file.name.lower().endswith(".xlsm")) else "filled_masterfile.xlsx"
        output_name = st.text_input("Output file name", value=default_name, help="Enter the name to use for the downloaded file (extension will be enforced).")

        # Process & Download
        can_process = (
            st.session_state.template_file is not None and
            not st.session_state.raw_df.empty and
            not st.session_state.mapping_table.empty
        )

        if not can_process:
            st.info("Please upload Template (Tab 1), Raw (Tab 2), and complete the mapping (Tab 3).")

        if st.button("⚙️ Process & Download", type="primary", disabled=not can_process):
            try:
                # Build writer mapping DataFrame from the grid (drop unmapped/invalid)
                mt = st.session_state.mapping_table.copy()
                mt = mt[(mt["Raw"].astype(str).str.strip() != "") & (mt["Template"].astype(str).str.strip() != "")]
                # Convert to writer format
                writer_map = mt[["Raw","Template"]].copy()
                writer_map.columns = ["raw_header","template_header"]

                if writer_map.empty:
                    st.error("No valid mapping rows selected.")
                else:
                    out_bytes = fill_first_sheet_by_headers(
                        template_bytes=BytesIO(st.session_state.template_file.getbuffer()),
                        mapping_df=writer_map,
                        raw_df=st.session_state.raw_df,
                        template_filename=st.session_state.template_file.name
                    )
                    keep_vba = st.session_state.template_file.name.lower().endswith(".xlsm")
                    final_name = sanitize_filename(output_name, keep_xlsm=keep_vba)
                    st.success("Done! Your updated masterfile is ready. Duplicate Partner SKU / Barcode cells are highlighted.")
                    st.download_button(
                        label="⬇️ Download Updated Masterfile",
                        data=out_bytes.getvalue(),
                        file_name=final_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            except Exception as e:
                st.exception(e)
