"""
matter_transform_app.py
------------------------
Streamlit UI for converting a raw client "matters" export (CSV or XLSX)
into the fixed migration-importer template CSV.

Run with:
    pip install streamlit pandas openpyxl --break-system-packages
    streamlit run matter_transform_app.py

Flow:
  1. Enter Org ID, upload Dropdown_<orgid>.json, upload Country Codes JSON,
     upload the raw CSV/XLSX file.
  2. Map every template column to a source column via dropdowns
     (required fields first, then optional).
  3. Map status / category / subcategory / country / type-of-mark values
     to the restricted options found in the JSON files (never guessed).
  4. Click "Save & Validate" -> all errors (duplicates, missing required
     fields, bad dates) are shown together, export blocked until fixed.
     Non-blocking warnings (e.g. missing subcategory) are shown too.
  5. On success, download the final CSV.
"""

import io
import re
import csv
import json
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------
# Constants
# ----------------------------------------------------------------------
# Fixed output template - only the headers originally supplied, no cf_* custom fields.
TEMPLATE_COLUMNS = [
    "bob", "mattercode", "country", "status", "category", "subcategory",
    "shorttitle", "longtitle", "user", "client", "client reference",
    "associate", "associate reference", "applicants", "inventors",
    "priority date", "priority number", "filing date", "filing country",
    "filing number", "publication date", "publication number",
    "grant date", "grant number", "expiry date", "hid_5 date", "notes",
]
NON_MAPPABLE_COLUMNS = {"bob"}
# Columns whose value comes from a restricted mapping step (dropdown JSON or
# country-codes JSON) rather than a raw 1:1 source-column copy.
VALUE_MAPPED_COLUMNS = {"status", "category", "subcategory", "country"}
REQUIRED_FIELDS = ["mattercode", "status", "category", "shorttitle", "user", "client", "country"]
DATE_COLUMNS = ["priority date", "filing date", "publication date", "grant date", "expiry date", "hid_5 date"]
SUBCATEGORY_GROUPS = [
    "MATTER_SUBCATEGORIES_COPYRIGHT",
    "MATTER_SUBCATEGORIES_DESIGN",
    "MATTER_SUBCATEGORIES_INVENTION",
    "MATTER_SUBCATEGORIES_OPPOSITION",
    "MATTER_SUBCATEGORIES_PATENT",
    "MATTER_SUBCATEGORIES_TM",
]
SUBCAT_REQUIRED_CATEGORIES = {
    "Copyright": "MATTER_SUBCATEGORIES_COPYRIGHT",
    "Design": "MATTER_SUBCATEGORIES_DESIGN",
    "Invention": "MATTER_SUBCATEGORIES_INVENTION",
    "Opposition": "MATTER_SUBCATEGORIES_OPPOSITION",
    "Patent": "MATTER_SUBCATEGORIES_PATENT",
    "Trade Mark": "MATTER_SUBCATEGORIES_TM",
}
SKIP = "-- Skip (leave blank) --"
CHOOSE = "-- Select a column --"


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------
def load_json_list(uploaded_file):
    """Handles both a bare JSON array and a dict wrapping one list value
    (e.g. {"select ...": [...]})."""
    raw = json.load(uploaded_file)
    if isinstance(raw, dict):
        for v in raw.values():
            if isinstance(v, list):
                return v
        raise ValueError("Unexpected JSON structure - no array found.")
    if isinstance(raw, list):
        return raw
    raise ValueError("Unexpected JSON structure.")


def get_descriptions_by_type(dropdown, type_names):
    if isinstance(type_names, str):
        type_names = [type_names]
    wanted = {t.strip().upper() for t in type_names}
    seen = []
    for entry in dropdown:
        t = str(entry.get("type", "")).strip().upper()
        if t in wanted:
            desc = str(entry.get("description", "")).strip()
            if desc and desc not in seen:
                seen.append(desc)
    return seen


def build_country_options(country_list):
    """Returns list of display strings 'CODE - Country Name', deduped, sorted by name."""
    seen = {}
    for entry in country_list:
        code = str(entry.get("code", "")).strip()
        name = str(entry.get("country", "")).strip()
        if code and name and code not in seen:
            seen[code] = name
    return [f"{code} - {name}" for code, name in sorted(seen.items(), key=lambda kv: kv[1])]


def code_from_option(option):
    """'US - United States' -> 'US'. Returns '' for skip/blank."""
    if not option or option == SKIP:
        return ""
    return option.split(" - ", 1)[0].strip()


def load_raw_file(uploaded_file):
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()
    if name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(io.BytesIO(data), dtype=str)
    elif name.endswith(".csv"):
        df = None
        last_err = None
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                df = pd.read_csv(io.BytesIO(data), dtype=str, keep_default_na=False, encoding=enc)
                break
            except (UnicodeDecodeError, UnicodeError) as e:
                last_err = e
                continue
        if df is None:
            raise ValueError(f"Could not decode CSV with any known encoding: {last_err}")
    else:
        raise ValueError("Unsupported file type. Use .csv or .xlsx")

    df = df.fillna("")
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
    return df


def parse_date_flexible(value):
    value = (value or "").strip()
    if value == "":
        return ""
    if re.match(r"^\d{4}-\d{2}-\d{2}$", value):
        return value
    formats = [
        "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%m-%d-%Y",
        "%d %b %Y", "%d %B %Y", "%b %d, %Y", "%B %d, %Y",
        "%Y/%m/%d", "%d.%m.%Y",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    try:
        serial = float(value)
        base = datetime(1899, 12, 30)
        return (base + timedelta(days=serial)).strftime("%Y-%m-%d")
    except (ValueError, OverflowError):
        pass
    raise ValueError(f"Unparsable date value: '{value}'")


def unique_values(df, col):
    if not col:
        return []
    return [v for v in dict.fromkeys(df[col].tolist())]


# ----------------------------------------------------------------------
# Page setup
# ----------------------------------------------------------------------
st.set_page_config(page_title="Matter Data Transformer", layout="wide")

st.markdown("""
<style>
.stamp-req {
    display:inline-block; padding:1px 8px; border:1px solid #B23A2E;
    color:#B23A2E; font-size:11px; font-weight:600; letter-spacing:0.06em;
    border-radius:3px; text-transform:uppercase;
}
.stamp-opt {
    display:inline-block; padding:1px 8px; border:1px solid #5B5F66;
    color:#5B5F66; font-size:11px; font-weight:600; letter-spacing:0.06em;
    border-radius:3px; text-transform:uppercase;
}
.section-label {
    font-size:12px; letter-spacing:0.12em; text-transform:uppercase;
    color:#5B5F66; font-weight:700; margin-bottom:2px;
}
</style>
""", unsafe_allow_html=True)

st.title("Matter Data Transformer")
st.caption("Raw client export -> migration-importer template CSV")

# ----------------------------------------------------------------------
# Session state init
# ----------------------------------------------------------------------
if "validated" not in st.session_state:
    st.session_state.validated = False
if "output_csv" not in st.session_state:
    st.session_state.output_csv = None
if "errors" not in st.session_state:
    st.session_state.errors = []
if "warnings" not in st.session_state:
    st.session_state.warnings = []

# ----------------------------------------------------------------------
# Step 1: Setup
# ----------------------------------------------------------------------
st.markdown('<div class="section-label">Step 1 - Setup</div>', unsafe_allow_html=True)
c0, c1, c2, c3 = st.columns([1, 2, 2, 2])
with c0:
    orgid = st.text_input("Org ID", placeholder="1251")
with c1:
    dropdown_file = st.file_uploader("Dropdown JSON (Dropdown_<orgid>.json)", type=["json"])
with c2:
    country_json_file = st.file_uploader("Country Codes JSON", type=["json"])
with c3:
    raw_file = st.file_uploader("Raw source file", type=["csv", "xlsx", "xls"])

if not (orgid and dropdown_file and country_json_file and raw_file):
    st.info("Enter the Org ID and upload all three files to continue.")
    st.stop()

try:
    dropdown = load_json_list(dropdown_file)
except Exception as e:
    st.error(f"Failed to parse dropdown JSON: {e}")
    st.stop()

try:
    country_list = load_json_list(country_json_file)
    country_options = build_country_options(country_list)
except Exception as e:
    st.error(f"Failed to parse country codes JSON: {e}")
    st.stop()

try:
    df = load_raw_file(raw_file)
except Exception as e:
    st.error(f"Failed to read raw file: {e}")
    st.stop()

st.success(
    f"Loaded {len(dropdown)} dropdown entries, {len(country_options)} country codes, "
    f"and {len(df)} raw rows ({len(df.columns)} columns)."
)
src_columns = list(df.columns)

with st.expander("Preview raw data (first 5 rows)"):
    st.dataframe(df.head())

st.divider()

# ----------------------------------------------------------------------
# Step 2: Column mapping - required fields first, then optional
# ----------------------------------------------------------------------
st.markdown('<div class="section-label">Step 2 - Map Columns</div>', unsafe_allow_html=True)

mappable_cols = [c for c in TEMPLATE_COLUMNS if c not in NON_MAPPABLE_COLUMNS and c not in VALUE_MAPPED_COLUMNS]
required_cols = [c for c in mappable_cols if c in REQUIRED_FIELDS]
optional_cols = [c for c in mappable_cols if c not in REQUIRED_FIELDS]

col_map = {}

st.markdown("**Required fields**")
rcols = st.columns(3)
for i, target_col in enumerate(required_cols):
    with rcols[i % 3]:
        st.markdown(f"`{target_col}` <span class='stamp-req'>required</span>", unsafe_allow_html=True)
        sel = st.selectbox(target_col, [CHOOSE] + src_columns, key=f"map_{target_col}", label_visibility="collapsed")
        col_map[target_col] = None if sel == CHOOSE else sel

# status / category / country are required but value-mapped, not 1:1 copies
st.markdown("**Required, value-mapped fields**")
vcol1, vcol2, vcol3 = st.columns(3)
with vcol1:
    st.markdown("`status` <span class='stamp-req'>required</span>", unsafe_allow_html=True)
    sel = st.selectbox("status", [CHOOSE] + src_columns, key="map_status_src", label_visibility="collapsed")
    status_src = None if sel == CHOOSE else sel
with vcol2:
    st.markdown("`category` <span class='stamp-req'>required</span>", unsafe_allow_html=True)
    sel = st.selectbox("category", [CHOOSE] + src_columns, key="map_category_src", label_visibility="collapsed")
    category_src = None if sel == CHOOSE else sel
with vcol3:
    st.markdown("`country` <span class='stamp-req'>required</span>", unsafe_allow_html=True)
    sel = st.selectbox("country", [CHOOSE] + src_columns, key="map_country_src", label_visibility="collapsed")
    country_src = None if sel == CHOOSE else sel

with st.expander("Optional fields", expanded=False):
    ocols = st.columns(3)
    for i, target_col in enumerate(optional_cols):
        with ocols[i % 3]:
            st.markdown(f"`{target_col}` <span class='stamp-opt'>optional</span>", unsafe_allow_html=True)
            sel = st.selectbox(target_col, [SKIP] + src_columns, key=f"map_{target_col}", label_visibility="collapsed")
            col_map[target_col] = None if sel == SKIP else sel

missing_required = (
    [c for c in required_cols if not col_map.get(c)]
    + (["status"] if not status_src else [])
    + (["category"] if not category_src else [])
    + (["country"] if not country_src else [])
)
if missing_required:
    st.warning(f"Still need a source column for: {', '.join(missing_required)}")
    st.stop()

st.divider()

# ----------------------------------------------------------------------
# Step 3: Dropdown-driven value mapping
# ----------------------------------------------------------------------
st.markdown('<div class="section-label">Step 3 - Map Values</div>', unsafe_allow_html=True)

# --- Status ---
status_options = get_descriptions_by_type(dropdown, ["MATTER_STATUS_LIVE", "MATTER_STATUS_CLOSED"])
status_value_map = {}
st.markdown("**Status values**")
if not status_options:
    st.error("No MATTER_STATUS_LIVE / MATTER_STATUS_CLOSED entries found in dropdown JSON.")
    st.stop()
status_unique = unique_values(df, status_src)
scols = st.columns(3)
for i, val in enumerate(status_unique):
    with scols[i % 3]:
        if val == "":
            status_value_map[val] = ""
            continue
        choice = st.selectbox(f"`{val}` ->", [SKIP] + status_options, key=f"statusmap_{val}")
        status_value_map[val] = "" if choice == SKIP else choice

# --- Category ---
category_options = get_descriptions_by_type(dropdown, "MATTER_CATEGORIES")
category_value_map = {}
st.markdown("**Category values**")
if not category_options:
    st.error("No MATTER_CATEGORIES entries found in dropdown JSON.")
    st.stop()
category_unique = unique_values(df, category_src)
ccols = st.columns(3)
for i, val in enumerate(category_unique):
    with ccols[i % 3]:
        if val == "":
            category_value_map[val] = ""
            continue
        choice = st.selectbox(f"`{val}` ->", [SKIP] + category_options, key=f"catmap_{val}")
        category_value_map[val] = "" if choice == SKIP else choice

# --- Country ---
country_value_map = {}
st.markdown("**Country values** *(mapped to country code)*")
if not country_options:
    st.error("No entries found in country codes JSON.")
    st.stop()
country_unique = unique_values(df, country_src)
ncols = st.columns(3)
for i, val in enumerate(country_unique):
    with ncols[i % 3]:
        if val == "":
            country_value_map[val] = ""
            continue
        choice = st.selectbox(f"`{val}` ->", [SKIP] + country_options, key=f"countrymap_{val}")
        country_value_map[val] = code_from_option(choice)

# --- Subcategory (optional) ---
st.markdown("**Subcategory (optional)**")
map_subcat = st.checkbox("Map a subcategory column?", value=True, key="do_subcat")
subcategory_src = None
subcategory_value_map = {}
if map_subcat:
    sc1, sc2 = st.columns(2)
    with sc1:
        sel = st.selectbox("Subcategory source column", [SKIP] + src_columns, key="map_subcategory_src")
        subcategory_src = None if sel == SKIP else sel
    with sc2:
        subcat_group = st.selectbox("Subcategory group to use", SUBCATEGORY_GROUPS, key="subcat_group")
    if subcategory_src:
        subcat_options = get_descriptions_by_type(dropdown, subcat_group)
        if not subcat_options:
            st.warning(f"No entries found in dropdown JSON for {subcat_group}. Subcategory will stay blank.")
        else:
            subcat_unique = unique_values(df, subcategory_src)
            scols2 = st.columns(3)
            for i, val in enumerate(subcat_unique):
                with scols2[i % 3]:
                    if val == "":
                        subcategory_value_map[val] = ""
                        continue
                    choice = st.selectbox(f"`{val}` ->", [SKIP] + subcat_options, key=f"subcatmap_{val}")
                    subcategory_value_map[val] = "" if choice == SKIP else choice

# --- Type of Mark (only if Trade Mark present) ---
trademark_present = any(
    category_value_map.get(v, "").strip().lower() == "trade mark" for v in df[category_src].tolist()
)
if trademark_present:
    st.markdown("**Type of Mark** *(category 'Trade Mark' detected)*")
    st.caption("Note: the current template has no output column for Type of Mark, "
               "so this mapping is captured for reference only and is not written to the export.")
    tom_col = st.selectbox("Type of Mark source column", [SKIP] + src_columns, key="map_tom_src")
    if tom_col != SKIP:
        tom_options = get_descriptions_by_type(dropdown, "TYPEOFMARK")
        if tom_options:
            tom_unique = unique_values(df, tom_col)
            tcols = st.columns(3)
            for i, val in enumerate(tom_unique):
                with tcols[i % 3]:
                    if val:
                        st.selectbox(f"`{val}` ->", [SKIP] + tom_options, key=f"tommap_{val}")
        else:
            st.warning("No TYPEOFMARK entries found in dropdown JSON.")

st.divider()

# ----------------------------------------------------------------------
# Step 4: Save & Validate
# ----------------------------------------------------------------------
st.markdown('<div class="section-label">Step 4 - Save & Validate</div>', unsafe_allow_html=True)

if st.button("Save & Validate", type="primary"):
    output_rows = []
    for _, row in df.iterrows():
        out = {}
        for target_col in TEMPLATE_COLUMNS:
            if target_col == "bob":
                out[target_col] = ""
            elif target_col == "status":
                out[target_col] = status_value_map.get(row[status_src], "")
            elif target_col == "category":
                out[target_col] = category_value_map.get(row[category_src], "")
            elif target_col == "country":
                out[target_col] = country_value_map.get(row[country_src], "")
            elif target_col == "subcategory":
                out[target_col] = subcategory_value_map.get(row[subcategory_src], "") if subcategory_src else ""
            else:
                src_col = col_map.get(target_col)
                out[target_col] = row[src_col] if src_col else ""
        output_rows.append(out)

    errors = []
    warnings = []
    seen_mattercodes = {}

    for idx, out in enumerate(output_rows, start=2):
        for field in REQUIRED_FIELDS:
            if not out.get(field, "").strip():
                errors.append(f"Row {idx}: missing required field '{field}'")

        mc = out.get("mattercode", "").strip()
        if mc:
            if mc in seen_mattercodes:
                errors.append(f"Row {idx}: duplicate mattercode '{mc}' (first seen at row {seen_mattercodes[mc]})")
            else:
                seen_mattercodes[mc] = idx

        for date_col in DATE_COLUMNS:
            val = out.get(date_col, "")
            if val == "":
                continue
            try:
                out[date_col] = parse_date_flexible(val)
            except ValueError as e:
                errors.append(f"Row {idx}: column '{date_col}' - {e}")

        cat = out.get("category", "").strip()
        if cat in SUBCAT_REQUIRED_CATEGORIES and not out.get("subcategory", "").strip():
            warnings.append(f"Row {idx}: category '{cat}' usually has a subcategory, but none was mapped. Please check.")

    st.session_state.errors = errors
    st.session_state.warnings = warnings
    st.session_state.validated = True

    if not errors:
        buf = io.StringIO()
        writer = csv.DictWriter(buf, fieldnames=TEMPLATE_COLUMNS)
        writer.writeheader()
        for out in output_rows:
            writer.writerow(out)
        st.session_state.output_csv = buf.getvalue()
    else:
        st.session_state.output_csv = None

if st.session_state.validated:
    if st.session_state.warnings:
        st.warning("Warnings (non-blocking):\n\n" + "\n".join(f"- {w}" for w in st.session_state.warnings))

    if st.session_state.errors:
        st.error(
            f"Export blocked - {len(st.session_state.errors)} error(s) found. Fix the source file and rerun.\n\n"
            + "\n".join(f"- {e}" for e in st.session_state.errors)
        )
    elif st.session_state.output_csv:
        st.success("All validations passed.")
        st.download_button(
            "Download output_matters.csv",
            data=st.session_state.output_csv,
            file_name="output_matters.csv",
            mime="text/csv",
        )
