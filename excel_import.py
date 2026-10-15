import io
import re
import pandas as pd
from database import get_connection

# NOTE: save_medicines now lives ONLY in database.py. A second copy used to live here and,
# because app.py star-imports both modules, it silently shadowed the safe version.

# ----------------------------------------------------------------------------
# TEMPLATE DEFINITION (column order shown to the user + used for download)
# ----------------------------------------------------------------------------
# header (as it appears in the sheet) -> (internal field, required?, help text)
TEMPLATE_BASE_COLUMNS = [
    ("ITEM DESCRIPTION", "medicine_name", True,  "Full item name exactly as you want it stored. Required."),
    ("BASE DRUG NAME",   "base_drug_name", False, "Optional grouping name. Leave blank to auto-detect from the item name."),
    ("STOCK ON HAND",    "stock_on_hand", False, "Optional number. Blank = 0 for new items, unchanged for existing ones."),
    ("IN TRANSIT",       "in_transit",    False, "Optional number."),
    ("AMC",              "amc",           False, "Optional number (average monthly consumption)."),
    ("REQUEST QTY",      "request_qty",   False, "Optional number."),
]
NUMERIC_FIELDS = ["stock_on_hand", "in_transit", "amc", "request_qty"]
RESERVED_FIELDS = {"medicine_name", "base_drug_name", "category", "import_status"} | set(NUMERIC_FIELDS)

DESCRIPTION_HEADERS = ["ITEM DESCRIPTION", "MEDICINE NAME", "ITEM NAME", "ITEM", "DESCRIPTION", "NAME"]


def _norm(header):
    return re.sub(r"\s+", " ", str(header)).strip().upper()


def extract_base_drug_name(full_name):
    """Isolates parent molecule definitions from raw input descriptions."""
    if pd.isna(full_name):
        return "UNKNOWN"

    text = str(full_name).strip()
    match = re.search(r'\d+', text)
    if match:
        base = text[:match.start()].strip()
        base = re.sub(r'[\s,\-]+$', '', base)
        if base:
            return base.upper()

    words = text.split()
    return words[0].strip(",;").upper() if words else "UNKNOWN"


def load_excel_sheets(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    return xl.sheet_names


def load_excel(uploaded_file, sheet_name=0):
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)


# ----------------------------------------------------------------------------
# TEMPLATE HELPERS
# ----------------------------------------------------------------------------
def get_template_columns(custom_cols=None):
    """Ordered list of column headers a sheet for this category should use."""
    return [h for h, _, _, _ in TEMPLATE_BASE_COLUMNS] + list(custom_cols or [])


def get_template_guide(custom_cols=None):
    """Small table describing each column (order, name, required, notes) for on-screen display."""
    rows = []
    for i, (header, _, required, note) in enumerate(TEMPLATE_BASE_COLUMNS, start=1):
        rows.append({"Order": i, "Column header": header, "Required": "Yes" if required else "No", "Notes": note})
    offset = len(TEMPLATE_BASE_COLUMNS)
    for j, col in enumerate(custom_cols or [], start=1):
        rows.append({"Order": offset + j, "Column header": col, "Required": "No", "Notes": "Custom column for this category."})
    return pd.DataFrame(rows)


def build_template_excel(category, custom_cols=None):
    """Builds a ready-to-fill .xlsx template (header row only) and returns its bytes."""
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Template"

    notes = {h: n for h, _, _, n in TEMPLATE_BASE_COLUMNS}
    headers = get_template_columns(custom_cols)
    fill = PatternFill("solid", start_color="1F4E78")
    for idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=idx, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.comment = Comment(notes.get(header, f"Custom column for {category}."), "Template")
        ws.column_dimensions[get_column_letter(idx)].width = 48 if idx == 1 else 18
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_blank_entry_df(custom_cols=None, rows=5):
    """Empty grid in the template column order, used for the 'type it in' option."""
    data = {}
    for header in get_template_columns(custom_cols):
        field = next((f for h, f, _, _ in TEMPLATE_BASE_COLUMNS if h == header), None)
        if field in NUMERIC_FIELDS:
            data[header] = pd.Series([None] * rows, dtype="float64")
        else:
            data[header] = pd.Series([None] * rows, dtype="object")
    return pd.DataFrame(data)


def parse_pasted_text(text):
    """
    Turns text pasted from Excel (tab separated) or a CSV into a DataFrame.
    * Tab separated  -> first line is the header row.
    * Comma header   -> used when the first line contains 'ITEM DESCRIPTION'.
    * Anything else  -> treated as a plain list: one item name per line.
    """
    text = (text or "").strip("\n\r ")
    if not text.strip():
        return pd.DataFrame()
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if "\t" in text:
        return pd.read_csv(io.StringIO("\n".join(lines)), sep="\t", dtype=str, keep_default_na=False)
    if "ITEM DESCRIPTION" in lines[0].upper() and "," in lines[0]:
        return pd.read_csv(io.StringIO("\n".join(lines)), sep=",", dtype=str, keep_default_na=False)
    return pd.DataFrame({"ITEM DESCRIPTION": [ln.strip() for ln in lines]})


# ----------------------------------------------------------------------------
# PREPARATION
# ----------------------------------------------------------------------------
def _find_description_column(df):
    norm_map = {_norm(c): c for c in df.columns}
    for wanted in DESCRIPTION_HEADERS:
        if wanted in norm_map:
            return norm_map[wanted]
    return next((c for c in df.columns if 'DESCRIPTION' in str(c).upper() or 'ITEM' in str(c).upper()), None)


def prepare_medicines(df, custom_cols=None):
    """
    Maps spreadsheet columns into the directory layout.

    Item names are kept EXACTLY as written in the sheet (only trimmed). Optional number columns
    and this category's custom columns are picked up only when their header matches the template.
    """
    custom_cols = custom_cols or []
    if df is None or df.empty:
        return pd.DataFrame()

    desc_col = _find_description_column(df)
    if desc_col is None:
        return pd.DataFrame()

    norm_map = {_norm(c): c for c in df.columns}
    out = pd.DataFrame({"medicine_name": df[desc_col]})

    # Optional explicit base drug name
    base_src = norm_map.get("BASE DRUG NAME")
    if base_src is not None and base_src != desc_col:
        out["_base_given"] = df[base_src]

    # Optional numeric template columns. Only read when the sheet is clearly in template form
    # (a legacy sheet that merely happens to contain an "AMC" column must not overwrite values).
    template_signature = {"STOCK ON HAND", "IN TRANSIT", "REQUEST QTY"}
    sheet_is_template = bool(template_signature & set(norm_map))
    for header, field, _, _ in TEMPLATE_BASE_COLUMNS:
        if field in NUMERIC_FIELDS and sheet_is_template:
            src = norm_map.get(header)
            if src is not None and src != desc_col:
                cleaned = df[src].astype(str).str.replace(",", "", regex=False).str.strip()
                out[field] = pd.to_numeric(cleaned, errors="coerce")

    # Category custom columns
    for col in custom_cols:
        src = norm_map.get(_norm(col))
        if src is not None and src != desc_col and col not in RESERVED_FIELDS:
            out[col] = df[src]

    out = out.dropna(subset=["medicine_name"])
    out["medicine_name"] = out["medicine_name"].astype(str).str.strip()
    out = out[out["medicine_name"] != '']
    out = out[~out["medicine_name"].str.contains("COMPATIBILITY REPORT|RUN ON|TOTAL", case=False, na=False)]
    out = out[out["medicine_name"].str.upper() != "NAN"]

    if out.empty:
        return pd.DataFrame()

    extracted = out["medicine_name"].apply(extract_base_drug_name)
    if "_base_given" in out.columns:
        given = out["_base_given"].fillna("").astype(str).str.strip()
        out["base_drug_name"] = [g.upper() if g and g.lower() != "nan" else e for g, e in zip(given, extracted)]
        out = out.drop(columns=["_base_given"])
    else:
        out["base_drug_name"] = extracted
    out = out[out["base_drug_name"] != "UNKNOWN"]

    # Same item twice (ignoring case) -> keep the first
    out = out[~out["medicine_name"].str.lower().duplicated()]

    # Tidy column order: names first, then numbers, then custom columns
    ordered = ["medicine_name", "base_drug_name"] + [f for f in NUMERIC_FIELDS if f in out.columns]
    ordered += [c for c in out.columns if c not in ordered]
    return out[ordered].reset_index(drop=True)
