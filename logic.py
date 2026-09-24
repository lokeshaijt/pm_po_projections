"""
Pure business logic for the PM PO Projection Mailer app.

Deliberately has no dependency on streamlit, so it can be unit-tested
and reused (e.g. from a CLI or notebook) independently of the UI.
"""

import html
import io
import re
from collections import defaultdict
from datetime import date

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Business-rule constants
# ---------------------------------------------------------------------------

TIERS = {
    "CFC": [250, 500, 1000, 3000, 5000, 10000],
    "CTN": [5000, 10000, 25000, 50000, 75000, 100000, 150000, 200000],
    "TRAY": [250, 500, 1000, 3000, 5000, 10000],
    "T-SHIRT": [1000],
    "AL-WIRE": [1000],
    "ALUMINIUM WIRE": [1000],
    "BOPP": [500],
    "BOPP FILM": [500],
    "ANGLE": [500],
    "SLIP SHEET": [250],
    "SLIP-SHEET": [250],
    "THREAD": [500],
}

# Supplier roster per Item Type. TRAY intentionally shares the CFC list.
SUPPLIERS_BY_TYPE = {
    "CFC": [
        "Abirami Packaging",
        "Arr Print Pack",
        "Mayura Packaging",
        "SRI RANGA INDUSTRIES",
        "Sri Venkateswara Packings",
    ],
    "TRAY": [
        "Abirami Packaging",
        "Arr Print Pack",
        "Mayura Packaging",
        "SRI RANGA INDUSTRIES",
        "Sri Venkateswara Packings",
    ],
    "CTN": [
        "Lovely Offset Printers (P) Ltd",
        "Salem Print Pack",
        "Shraddha Saburi Printers Private Limited",
    ],
}

ITEM_TYPES = ["CFC", "CTN", "TRAY"]

# Known corrections for items whose live description doesn't exactly match
# the master (e.g. "SC ENV" vs "DC ENV" naming variants). Extend as needed.
CATEGORY_OVERRIDES: dict = {
    # 'ITM09821': '25TB DC ENV (AFRICA)',  # example — add real codes/categories here
}


# ---------------------------------------------------------------------------
# Parsing helpers
# ---------------------------------------------------------------------------

def normalize_header(text) -> str:
    return re.sub(r"\s+", " ", (text or "").__str__().strip().upper())


def normalize_name(text) -> str:
    if text is None:
        return ""
    s = str(text).replace("\xa0", " ")
    return re.sub(r"\s+", " ", s).strip().upper()


def find_header_row(ws, required_header_text: str, max_scan_rows: int = 12):
    """Scan the first rows of a worksheet for a row containing the given
    header text; return (row_number, {normalized_header: col_number})."""
    target = normalize_header(required_header_text)
    for r in range(1, min(max_scan_rows, ws.max_row) + 1):
        row_cells = list(ws[r])
        found = any(normalize_header(c.value) == target for c in row_cells)
        if found:
            col_map = {}
            for c in row_cells:
                if c.value is not None:
                    col_map[normalize_header(c.value)] = c.column
            return r, col_map
    return None, None


def col(col_map, name):
    return col_map.get(normalize_header(name))


def load_category_master(file) -> tuple:
    """Load the Item Category Master workbook (ITEM NAME, Item Type,
    ITEM CATEGORY columns, in whichever order they appear)."""
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb["Item Category"] if "Item Category" in wb.sheetnames else wb.worksheets[0]
    header_row, col_map = find_header_row(ws, "ITEM NAME", max_scan_rows=3)
    if header_row is None:
        header_row, col_map = 1, {
            normalize_header(c.value): c.column for c in ws[1] if c.value
        }
    c_name = col(col_map, "ITEM NAME")
    c_type = col(col_map, "Item Type")
    c_cat = col(col_map, "ITEM CATEGORY")

    name_to_cat, name_to_type = {}, {}
    for r in range(header_row + 1, ws.max_row + 1):
        name = ws.cell(row=r, column=c_name).value if c_name else None
        if not name:
            continue
        norm = normalize_name(name)
        if c_cat:
            cat = ws.cell(row=r, column=c_cat).value
            if cat:
                name_to_cat[norm] = str(cat).strip()
        if c_type:
            itype = ws.cell(row=r, column=c_type).value
            if itype:
                name_to_type[norm] = str(itype).strip().upper()
    return name_to_cat, name_to_type


def parse_requirement_workbook(file):
    """Parse the PM Requirement workbook's Shortfall + Laminates sheets.

    Returns a dict with: item_names, categories (coarse), week_shortfalls
    (per item, all 8 weeks in the current window), all_weeks (list[int]),
    projection_weeks (last 6 of the 8), laminated_codes.
    """
    wb = openpyxl.load_workbook(file, data_only=True)

    short_ws = None
    short_header = None
    for ws in wb.worksheets:
        if normalize_header(ws.title) == "SHORTFALL":
            r, cm = find_header_row(ws, "Shortfall Quantity")
            if r:
                short_ws, short_header = ws, (r, cm)
                break
    if short_ws is None:
        for ws in wb.worksheets:
            r, cm = find_header_row(ws, "Shortfall Quantity")
            if r:
                short_ws, short_header = ws, (r, cm)
                break
    if short_ws is None:
        raise ValueError("Could not find a 'Shortfall Quantity' column in the Requirement file.")

    header_row, col_map = short_header
    c_code = col(col_map, "Nav Item Code")
    c_name = col(col_map, "Item Name")
    c_cat = col(col_map, "Item Category")
    c_type = col(col_map, "Item Type")

    # Discover which weeks appear under the "Shortfall Weeks" group, in order.
    week_cols = []
    for cnum in range(1, short_ws.max_column + 1):
        header_val = short_ws.cell(row=header_row, column=cnum).value
        h = normalize_header(header_val)
        m = re.match(r"WEEK (\d+)", h)
        if m and cnum > (c_cat or 0):
            week_cols.append((int(m.group(1)), cnum))
        if len(week_cols) >= 8:
            break
    week_cols = week_cols[:8]
    all_weeks = [w for w, _ in week_cols]
    week_colmap = dict(week_cols)

    item_names, categories, item_types = {}, {}, {}
    week_shortfalls = {}
    for r in range(header_row + 1, short_ws.max_row + 1):
        code = short_ws.cell(row=r, column=c_code).value if c_code else None
        if not code or str(code).strip() in ("", "-"):
            continue
        code = str(code).strip()
        if c_name:
            item_names[code] = short_ws.cell(row=r, column=c_name).value
        if c_cat:
            categories[code] = short_ws.cell(row=r, column=c_cat).value
        if c_type:
            item_types[code] = short_ws.cell(row=r, column=c_type).value
        wk = {}
        for w, cnum in week_colmap.items():
            v = short_ws.cell(row=r, column=cnum).value
            wk[w] = v if v is not None else 0
        week_shortfalls[code] = wk

    laminated_codes = set()
    for ws in wb.worksheets:
        if "LAMIN" not in normalize_header(ws.title):
            continue
        r, cm = find_header_row(ws, "Nav Item Code")
        if not r:
            continue
        c_code2 = col(cm, "Nav Item Code")
        for rr in range(r + 1, ws.max_row + 1):
            v = ws.cell(row=rr, column=c_code2).value
            if v:
                laminated_codes.add(str(v).strip())
        break

    projection_weeks = all_weeks[2:] if len(all_weeks) >= 3 else all_weeks

    return {
        "item_names": item_names,
        "categories": categories,
        "week_shortfalls": week_shortfalls,
        "all_weeks": all_weeks,
        "projection_weeks": projection_weeks,
        "laminated_codes": laminated_codes,
    }


def parse_pending_po(file) -> pd.DataFrame:
    """Load the Pending PO export (Sheet2, or whichever sheet has
    'Outstanding Quantity'), dedupe exact-duplicate rows, and normalize
    the columns this app needs."""
    wb = openpyxl.load_workbook(file, data_only=True)
    target_ws = None
    header_row, col_map = None, None
    for ws in wb.worksheets:
        r, cm = find_header_row(ws, "Outstanding Quantity")
        if r:
            target_ws, header_row, col_map = ws, r, cm
            break
    if target_ws is None:
        raise ValueError("Could not find an 'Outstanding Quantity' column in the Pending PO file.")

    c_doc = col(col_map, "Document No")
    c_odate = col(col_map, "Order Date")
    c_cust = col(col_map, "Customer Name")
    c_item = col(col_map, "Item No")
    c_desc = col(col_map, "Item Description")
    c_deliv = col(col_map, "Delivery Date")
    c_qty = col(col_map, "Quantity")
    c_outstanding = col(col_map, "Outstanding Quantity")

    rows = []
    for r in range(header_row + 1, target_ws.max_row + 1):
        item_no = target_ws.cell(row=r, column=c_item).value if c_item else None
        if not item_no:
            continue
        rows.append(
            {
                "Document No": target_ws.cell(row=r, column=c_doc).value,
                "Order Date": target_ws.cell(row=r, column=c_odate).value,
                "Customer Name": target_ws.cell(row=r, column=c_cust).value,
                "Item No": str(item_no).strip(),
                "Item Description": target_ws.cell(row=r, column=c_desc).value,
                "Delivery Date": target_ws.cell(row=r, column=c_deliv).value,
                "Quantity": target_ws.cell(row=r, column=c_qty).value,
                "Outstanding Quantity": target_ws.cell(row=r, column=c_outstanding).value,
            }
        )
    df = pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)
    df["Order Date"] = pd.to_datetime(df["Order Date"], dayfirst=True, errors="coerce")
    df["Delivery Date"] = pd.to_datetime(df["Delivery Date"], dayfirst=True, errors="coerce")
    return df


# ---------------------------------------------------------------------------
# Business logic
# ---------------------------------------------------------------------------

def compute_moq(item_code, category, qty, item_names, laminated_codes):
    if qty is None or qty <= 0:
        return None
    cat_norm = (category or "").strip().upper()
    name_up = (item_names.get(item_code) or "").upper()
    if item_code in laminated_codes:
        tiers = [300]
    elif cat_norm == "POUCH" and "GARANT" in name_up:
        tiers = [5000, 10000, 25000, 50000]
    else:
        tiers = TIERS.get(cat_norm)
    if not tiers:
        return None
    for t in tiers:
        if qty <= t:
            return t
    return round(qty * 1.02)


def build_item_category_map(req_data, master_name_to_cat, master_name_to_type):
    """Map each Nav Item Code to (fine category, item type)."""
    item_names = req_data["item_names"]
    categories = req_data["categories"]
    item_to_cat, item_to_type = {}, {}
    for code, name in item_names.items():
        norm = normalize_name(name)
        fine_cat = master_name_to_cat.get(norm)
        if code in CATEGORY_OVERRIDES:
            fine_cat = CATEGORY_OVERRIDES[code]
        if fine_cat:
            item_to_cat[code] = fine_cat
            item_to_type[code] = master_name_to_type.get(norm, (categories.get(code) or "").strip().upper())
    return item_to_cat, item_to_type


def build_item_week_projections(req_data):
    item_names = req_data["item_names"]
    categories = req_data["categories"]
    laminated_codes = req_data["laminated_codes"]
    week_shortfalls = req_data["week_shortfalls"]
    projection_weeks = req_data["projection_weeks"]

    item_week_proj = {}
    for code, wk in week_shortfalls.items():
        cat = categories.get(code)
        item_week_proj[code] = {
            w: compute_moq(code, cat, wk.get(w, 0), item_names, laminated_codes)
            for w in projection_weeks
        }
    return item_week_proj


def build_po_issued_df(po_df: pd.DataFrame, today: date, window_days: int) -> pd.DataFrame:
    today_ts = pd.Timestamp(today)
    cutoff = today_ts + pd.Timedelta(days=window_days)
    window = po_df[po_df["Delivery Date"] <= cutoff].copy()
    window["No. of days to arrive"] = (window["Delivery Date"] - today_ts).dt.days
    window = window.sort_values(["Customer Name", "Item No", "Delivery Date"]).reset_index(drop=True)
    return window


def build_category_table(item_week_proj, item_to_cat, item_to_type, category_order, projection_weeks):
    """One row per category (in category_order), Item Type, week columns, Total."""
    cat_week_sum = defaultdict(lambda: {w: 0.0 for w in projection_weeks})
    cat_type = {}
    for code, projs in item_week_proj.items():
        cat = item_to_cat.get(code)
        if not cat:
            continue
        cat_type[cat] = item_to_type.get(code, cat_type.get(cat, ""))
        for w in projection_weeks:
            v = projs.get(w)
            if v:
                cat_week_sum[cat][w] += v

    rows = []
    for cat in category_order:
        weeksum = cat_week_sum.get(cat, {w: 0.0 for w in projection_weeks})
        total = sum(weeksum.values())
        row = {"Item Category": cat, "Item Type": cat_type.get(cat, "")}
        for w in projection_weeks:
            row[f"Wk #{w}"] = weeksum[w]
        row["Total"] = total
        rows.append(row)
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Excel report generation
# ---------------------------------------------------------------------------

def style_header_cell(cell, fill_hex="F4B183"):
    thin = Side(style="thin", color="000000")
    cell.font = Font(name="Arial", bold=True, size=11)
    cell.fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def style_data_cell(cell, bold=False, color=None, fill_hex=None):
    thin = Side(style="thin", color="000000")
    cell.font = Font(name="Arial", size=11, bold=bold, color=color)
    if fill_hex:
        cell.fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def generate_report_workbook(po_issued: pd.DataFrame, category_df: pd.DataFrame, projection_weeks) -> bytes:
    wb = openpyxl.Workbook()

    # ---- PO Issued ----
    ws1 = wb.active
    ws1.title = "PO Issued"
    headers1 = ["Supplier Name", "PO number", "PO date", "Item Description",
                "PO Qty", "Outstanding Qty", "Delivery date", "No. of days to arrive"]
    ws1.append(headers1)
    for cidx in range(1, len(headers1) + 1):
        style_header_cell(ws1.cell(row=1, column=cidx))

    for i, row in po_issued.iterrows():
        r = i + 2
        ws1.cell(row=r, column=1, value=row["Customer Name"])
        style_data_cell(ws1.cell(row=r, column=1))
        ws1.cell(row=r, column=1).alignment = Alignment(vertical="center")
        ws1.cell(row=r, column=2, value=row["Document No"])
        style_data_cell(ws1.cell(row=r, column=2))
        c = ws1.cell(row=r, column=3, value=row["Order Date"])
        style_data_cell(c)
        c.number_format = "DD/MM/YYYY"
        ws1.cell(row=r, column=4, value=row["Item Description"])
        style_data_cell(ws1.cell(row=r, column=4))
        ws1.cell(row=r, column=4).alignment = Alignment(vertical="center")
        c = ws1.cell(row=r, column=5, value=float(row["Quantity"]) if pd.notna(row["Quantity"]) else None)
        style_data_cell(c)
        c.number_format = "#,##0"
        c = ws1.cell(row=r, column=6, value=float(row["Outstanding Quantity"]) if pd.notna(row["Outstanding Quantity"]) else None)
        style_data_cell(c)
        c.number_format = "#,##0"
        c = ws1.cell(row=r, column=7, value=row["Delivery Date"])
        style_data_cell(c)
        c.number_format = "DD/MM/YYYY"
        days = int(row["No. of days to arrive"])
        delayed = days < 0
        c = ws1.cell(row=r, column=8, value=days)
        style_data_cell(c, bold=delayed, color="FFFF0000" if delayed else None,
                         fill_hex="FCE5E5" if delayed else None)

    widths1 = {"A": 30, "B": 20, "C": 12, "D": 42, "E": 10, "F": 12, "G": 13, "H": 12}
    for c, w in widths1.items():
        ws1.column_dimensions[c].width = w
    ws1.row_dimensions[1].height = 30
    last1 = ws1.max_row
    ws1.auto_filter.ref = f"A1:H{last1}"
    ws1.freeze_panes = "A2"
    ws1.page_setup.orientation = "landscape"
    ws1.page_setup.fitToWidth = 1
    ws1.page_setup.fitToHeight = 0
    ws1.sheet_properties.pageSetUpPr.fitToPage = True

    # ---- PO Projection (only categories with a projection) ----
    ws2 = wb.create_sheet("PO Projection")
    nonzero_df = category_df[category_df["Total"] > 0].reset_index(drop=True)

    headers2 = ["Item Category", "Item Type"] + [f"Wk #{w}" for w in projection_weeks] + ["Total"]
    ws2.append(headers2)
    for cidx in range(1, len(headers2) + 1):
        style_header_cell(ws2.cell(row=1, column=cidx))

    for i, row in nonzero_df.iterrows():
        r = i + 2
        ws2.cell(row=r, column=1, value=row["Item Category"])
        style_data_cell(ws2.cell(row=r, column=1))
        ws2.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center")
        ws2.cell(row=r, column=2, value=row["Item Type"])
        style_data_cell(ws2.cell(row=r, column=2))
        for wi, w in enumerate(projection_weeks):
            val = row[f"Wk #{w}"]
            c = ws2.cell(row=r, column=3 + wi)
            if val and val > 0:
                c.value = float(val)
                c.number_format = "#,##0"
            else:
                c.value = "-"
            style_data_cell(c)
        c = ws2.cell(row=r, column=3 + len(projection_weeks), value=float(row["Total"]))
        c.number_format = "#,##0"
        style_data_cell(c, bold=True)

    widths2 = {"A": 32, "B": 12}
    for c, w in widths2.items():
        ws2.column_dimensions[c].width = w
    for i in range(len(projection_weeks) + 1):
        ws2.column_dimensions[get_column_letter(3 + i)].width = 13
    ws2.row_dimensions[1].height = 24
    last2 = ws2.max_row
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(headers2))}{last2}"
    ws2.freeze_panes = "A2"
    ws2.page_setup.orientation = "landscape"
    ws2.page_setup.fitToWidth = 1
    ws2.page_setup.fitToHeight = 0
    ws2.sheet_properties.pageSetUpPr.fitToPage = True

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Email
# ---------------------------------------------------------------------------

EMAIL_TEMPLATE_INTRO = (
    "Please find the Item category wise projection Qty for upcoming weeks.\n\n"
    "Note: Projection Qty is only to Procure raw materials. Upon the receipt "
    "of PO, you can print as per the delivery schedule."
)


# ---------------------------------------------------------------------------
# Supplier allocation (per category/week cell)
# ---------------------------------------------------------------------------

def allocate_to_suppliers(category_df: pd.DataFrame, selections: dict, projection_weeks) -> pd.DataFrame:
    """Split each category/week projection evenly across the suppliers picked
    for that cell. `selections` maps (item_category, week) -> [supplier, ...].

    Returns one row per (Supplier, Item Category) with the allocated quantity
    in each "Wk #N" column, ordered by supplier, then Item Type, then category.
    """
    week_cols = [f"Wk #{w}" for w in projection_weeks]
    by_cat = category_df.set_index("Item Category")
    alloc = defaultdict(lambda: {c: 0.0 for c in week_cols})
    cat_types = {}
    for (cat, w), suppliers in selections.items():
        suppliers = [s for s in dict.fromkeys(suppliers or []) if s]
        if not suppliers or cat not in by_cat.index:
            continue
        qty = by_cat.at[cat, f"Wk #{w}"]
        if not qty or qty <= 0:
            continue
        share = qty / len(suppliers)
        for s in suppliers:
            alloc[(s, cat)][f"Wk #{w}"] += share
            cat_types[cat] = by_cat.at[cat, "Item Type"]

    columns = ["Supplier", "Item Category", "Item Type"] + week_cols
    if not alloc:
        return pd.DataFrame(columns=columns)

    supplier_order = list(dict.fromkeys(s for sups in SUPPLIERS_BY_TYPE.values() for s in sups))
    type_order = {t: i for i, t in enumerate(ITEM_TYPES)}

    def sort_key(key):
        s, cat = key
        s_idx = supplier_order.index(s) if s in supplier_order else len(supplier_order)
        return (s_idx, s, type_order.get(cat_types[cat], len(type_order)), cat)

    rows = []
    for key in sorted(alloc, key=sort_key):
        s, cat = key
        row = {"Supplier": s, "Item Category": cat, "Item Type": cat_types[cat]}
        row.update({c: round(v) for c, v in alloc[key].items()})
        rows.append(row)
    return pd.DataFrame(rows, columns=columns)


def week_label(w) -> str:
    return f"Wk {w}"


def build_allocation_html(alloc_df: pd.DataFrame, projection_weeks) -> str:
    """Render the supplier allocation as a bordered HTML table: supplier name
    merged down its rows, then Item category, Item type, and one column per
    week (blank where nothing is allocated). Inline styles so it survives
    email clients."""
    cell = "border:1px solid #000;padding:4px 8px;font-family:Calibri,Arial,sans-serif;font-size:14px;"
    head = ["Item category", "Item type"] + [week_label(w) for w in projection_weeks]
    out = ['<table style="border-collapse:collapse;">', "<tr>", f'<th style="{cell}"></th>']
    out += [f'<th style="{cell}text-align:left;font-weight:normal;">{h}</th>' for h in head]
    out.append("</tr>")
    for supplier, grp in alloc_df.groupby("Supplier", sort=False):
        for i, (_, row) in enumerate(grp.iterrows()):
            out.append("<tr>")
            if i == 0:
                out.append(
                    f'<td rowspan="{len(grp)}" style="{cell}text-align:center;vertical-align:middle;">'
                    f"{html.escape(supplier)}</td>"
                )
            out.append(f'<td style="{cell}">{html.escape(str(row["Item Category"]))}</td>')
            out.append(f'<td style="{cell}text-align:center;">{html.escape(str(row["Item Type"]))}</td>')
            for w in projection_weeks:
                v = row[f"Wk #{w}"]
                out.append(f'<td style="{cell}text-align:right;">{f"{v:.0f}" if v else ""}</td>')
            out.append("</tr>")
    out.append("</table>")
    return "".join(out)


def build_supplier_email_html(supplier_name: str, supplier_alloc_df: pd.DataFrame, projection_weeks) -> str:
    intro = html.escape(EMAIL_TEMPLATE_INTRO).replace("\n", "<br>")
    return (
        '<div style="font-family:Calibri,Arial,sans-serif;font-size:14px;">'
        f"<p>Hello {html.escape(supplier_name)},</p>"
        f"<p>{intro}</p>"
        f"{build_allocation_html(supplier_alloc_df, projection_weeks)}"
        "<p>Thank you.</p></div>"
    )


def build_supplier_allocation_text(supplier_name: str, supplier_alloc_df: pd.DataFrame, projection_weeks) -> str:
    """Plain-text fallback of the HTML email, for clients that don't render HTML."""
    lines = [f"Hello {supplier_name},", "", EMAIL_TEMPLATE_INTRO, ""]
    for _, row in supplier_alloc_df.iterrows():
        lines.append(f"- {row['Item Category']} ({row['Item Type']})")
        for w in projection_weeks:
            v = row[f"Wk #{w}"]
            if v:
                lines.append(f"    {week_label(w)}: {v:.0f}")
    lines += ["", "Thank you."]
    return "\n".join(lines)
