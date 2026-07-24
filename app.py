"""
GSH Rental Balance Qty — Web App
=================================
Upload a ZIP of rental bill PDFs -> get a formatted Excel back.

Same logic as the "rental-combined" skill, with ONE change Rahul asked for:
NEGATIVE quantities are never added into any sum. When the same party/item
appears across multiple bills, only positive quantities are added together;
negative lines are simply skipped (not subtracted, not summed).

Run locally:
    pip install -r requirements.txt
    streamlit run app.py

Deploy for free (shareable link):
    Push this folder to a GitHub repo, then deploy on share.streamlit.io
    (see deployment steps Rahul was given in chat).
"""
import os
import re
import glob
import shutil
import tempfile
import zipfile
from collections import defaultdict

import streamlit as st
from pdfminer.high_level import extract_text
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

st.set_page_config(page_title="GSH Rental Balance Qty", page_icon="📊", layout="centered")

# ════════════════════════════════════════════════════════
#  PARSING  (same as parse_all.py)
# ════════════════════════════════════════════════════════

LOCATION_MAP = {
    "GSHP": "GSHP Pune", "PUNE": "GSHP Pune",
    "GSHG": "GSHG Gurugram", "GURUGRAM": "GSHG Gurugram",
    "KSH": "KSH Karnawati", "KARNAWATI": "KSH Karnawati",
    "SSS": "SSS Ankleshwar", "ANKLESHWAR": "SSS Ankleshwar", "SUMIT": "SSS Ankleshwar",
}

def clean_location(folder_name):
    upper = folder_name.upper()
    for key, val in LOCATION_MAP.items():
        if key in upper:
            return val
    name = re.sub(r'[\-_]+', ' ', folder_name).strip()
    return name[:28]

def extract_balance_qty(text):
    text = re.sub(r'\s+', ' ', text)
    match = re.search(
        r'Balance\s+Qty\s*[:\-]?\s*(.*?)(?:Cartage|Total\s+Amount|Grand\s+Total|$)',
        text, re.IGNORECASE | re.DOTALL
    )
    if not match:
        return {}
    balance_text = match.group(1).strip()
    items = {}
    for part in balance_text.split(','):
        part = part.strip()
        m = re.match(r'^(.+?)\s*-\s*([\d\.\-]+)\s*$', part)
        if m:
            item = m.group(1).strip().upper()
            try:
                qty = float(m.group(2))
                if item:
                    items[item] = items.get(item, 0) + qty
            except ValueError:
                pass
    return items

def get_party(filepath):
    base = os.path.basename(filepath).replace('.pdf', '')
    parts = base.split('_')
    if len(parts) >= 2:
        return parts[1].strip()
    return base

def parse_folder(root, progress_cb=None):
    bills, all_items, no_balance, errors = [], set(), [], []
    pdf_files = glob.glob(os.path.join(root, '**', '*.pdf'), recursive=True)
    if not pdf_files:
        pdf_files = glob.glob(os.path.join(root, '*.pdf'))

    for idx, pdf_path in enumerate(sorted(pdf_files)):
        if progress_cb:
            progress_cb(idx + 1, len(pdf_files), os.path.basename(pdf_path))
        parent_dir = os.path.dirname(pdf_path)
        parent_name = os.path.basename(parent_dir)
        if os.path.normpath(parent_dir) == os.path.normpath(root):
            location = clean_location(os.path.basename(root))
        else:
            location = clean_location(parent_name)
        party = get_party(pdf_path)
        try:
            text = extract_text(pdf_path)
            items = extract_balance_qty(text)
            if not items:
                no_balance.append(os.path.basename(pdf_path))
            bills.append({
                "filename": os.path.basename(pdf_path),
                "location": location, "party": party, "items": items
            })
            all_items.update(items.keys())
        except Exception as e:
            errors.append(f"{os.path.basename(pdf_path)}: {e}")

    return {
        "bills": bills,
        "locations": sorted(set(b['location'] for b in bills)),
        "parties": sorted(set(b['party'] for b in bills)),
        "all_items": sorted(all_items),
        "no_balance_count": len(no_balance),
        "errors": errors,
        "pdf_count": len(pdf_files),
    }


# ════════════════════════════════════════════════════════
#  EXCEL BUILDING  (same as build_combined_excel.py, but
#  negative quantities are SKIPPED when combining bills —
#  never added into any party/location/dashboard total)
# ════════════════════════════════════════════════════════

DARK_BLUE, MED_BLUE = "1F4E78", "2F75B5"
GREEN_DARK, GREEN_MED = "375623", "538135"
WHITE, ALT_GRAY = "FFFFFF", "F2F2F2"

def mk_fill(hex_color):
    return PatternFill(fill_type="solid", fgColor=hex_color)

def thin_border():
    s = Side(style='thin', color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

CTR = Alignment(horizontal="center", vertical="center")
LFT = Alignment(horizontal="left", vertical="center")
WRP = Alignment(horizontal="center", vertical="center", wrap_text=True)

def title_row(ws, row, text, ncols, dark_hex, height=28):
    ws.row_dimensions[row].height = height
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name="Arial", bold=True, size=13, color=WHITE)
    c.fill = mk_fill(dark_hex); c.alignment = CTR

def info_row(ws, row, text, ncols, med_hex, height=18):
    ws.row_dimensions[row].height = height
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name="Arial", size=9, italic=True, color=WHITE)
    c.fill = mk_fill(med_hex); c.alignment = CTR

def header_cell(ws, row, col, val, dark_hex, algn=CTR):
    c = ws.cell(row=row, column=col, value=val)
    c.font = Font(name="Arial", bold=True, size=10, color=WHITE)
    c.fill = mk_fill(dark_hex); c.alignment = algn; c.border = thin_border()
    return c

def data_cell(ws, row, col, val, row_fill, algn=CTR, fmt=None):
    c = ws.cell(row=row, column=col, value=val)
    c.font = Font(name="Arial", size=10); c.fill = row_fill
    c.alignment = algn; c.border = thin_border()
    if fmt:
        c.number_format = fmt
    return c

def build_dashboard(wb, sheet_title, title_text, top_parties, top_items,
                     n_parties, n_items, grand_total, dark_hex, med_hex, bar_color):
    ws = wb.create_sheet(title=sheet_title)
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 32
    ws.merge_cells("A1:R1")
    c = ws.cell(row=1, column=1, value=title_text)
    c.font = Font(name="Arial", bold=True, size=14, color=WHITE)
    c.fill = mk_fill(dark_hex); c.alignment = CTR

    ws.row_dimensions[2].height = 20
    ws.merge_cells("A2:R2")
    c = ws.cell(row=2, column=1,
        value=f"Parties: {n_parties}   |   Item Types: {n_items}   |   Grand Total Qty: {int(grand_total):,}")
    c.font = Font(name="Arial", size=10, italic=True, color=WHITE)
    c.fill = mk_fill(med_hex); c.alignment = CTR

    ws.cell(row=3, column=1, value="Party").font = Font(bold=True, name="Arial", size=9)
    ws.cell(row=3, column=2, value="Total Qty").font = Font(bold=True, name="Arial", size=9)
    ws.cell(row=3, column=4, value="Item").font = Font(bold=True, name="Arial", size=9)
    ws.cell(row=3, column=5, value="Total Qty").font = Font(bold=True, name="Arial", size=9)

    for i, (party, qty) in enumerate(top_parties, start=4):
        ws.cell(row=i, column=1, value=party).font = Font(name="Arial", size=9)
        c = ws.cell(row=i, column=2, value=int(qty)); c.font = Font(name="Arial", size=9)
        c.number_format = '#,##0'
    for i, (item, qty) in enumerate(top_items, start=4):
        ws.cell(row=i, column=4, value=item).font = Font(name="Arial", size=9)
        c = ws.cell(row=i, column=5, value=int(qty)); c.font = Font(name="Arial", size=9)
        c.number_format = '#,##0'

    n_p = len(top_parties)
    chart1 = BarChart()
    chart1.type = "bar"; chart1.title = "Top 20 Parties — Total Qty"
    chart1.y_axis.title = "Party"; chart1.x_axis.title = "Total Qty"
    chart1.style = 10; chart1.width = 22; chart1.height = 14
    chart1.add_data(Reference(ws, min_col=2, min_row=3, max_row=3+n_p), titles_from_data=True)
    chart1.set_categories(Reference(ws, min_col=1, min_row=4, max_row=3+n_p))
    chart1.series[0].graphicalProperties.solidFill = bar_color
    ws.add_chart(chart1, "A25")

    n_i = len(top_items)
    chart2 = BarChart()
    chart2.type = "bar"; chart2.title = "Top 20 Items — Total Qty (All Parties)"
    chart2.y_axis.title = "Item"; chart2.x_axis.title = "Total Qty"
    chart2.style = 10; chart2.width = 22; chart2.height = 14
    chart2.add_data(Reference(ws, min_col=5, min_row=3, max_row=3+n_i), titles_from_data=True)
    chart2.set_categories(Reference(ws, min_col=4, min_row=4, max_row=3+n_i))
    chart2.series[0].graphicalProperties.solidFill = dark_hex
    ws.add_chart(chart2, "L25")

    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 2
    ws.column_dimensions["D"].width = 36
    ws.column_dimensions["E"].width = 14

def build_party_matrix(wb, sheet_title, title_text, party_data,
                        sorted_parties, all_items, dark_hex, med_hex):
    ws = wb.create_sheet(title=sheet_title)
    total_cols = 2 + len(all_items)
    title_row(ws, 1, title_text, total_cols, dark_hex)
    info_row(ws, 2,
        f"Parties: {len(sorted_parties)}   |   Items: {len(all_items)}   |   Negative qty never added into totals",
        total_cols, med_hex)

    ws.row_dimensions[3].height = 50
    header_cell(ws, 3, 1, "S.No", dark_hex, CTR)
    header_cell(ws, 3, 2, "PARTY NAME", dark_hex, Alignment(horizontal="left", vertical="center", wrap_text=True))
    for idx, item in enumerate(all_items, start=3):
        c = ws.cell(row=3, column=idx, value=item)
        c.font = Font(name="Arial", bold=True, size=9, color=WHITE)
        c.fill = mk_fill(med_hex); c.alignment = WRP; c.border = thin_border()

    for row_num, party in enumerate(sorted_parties, start=1):
        er = 3 + row_num
        ws.row_dimensions[er].height = 16
        row_fill = mk_fill(ALT_GRAY) if row_num % 2 == 0 else mk_fill(WHITE)
        data_cell(ws, er, 1, row_num, row_fill, CTR)
        data_cell(ws, er, 2, party, row_fill, LFT)
        p_items = party_data.get(party, {})
        for col_idx, item in enumerate(all_items, start=3):
            qty = p_items.get(item, None)
            data_cell(ws, er, col_idx, qty, row_fill, CTR, '#,##0' if qty else None)

    total_row = 3 + len(sorted_parties) + 1
    data_start, data_end = 4, 3 + len(sorted_parties)
    ws.row_dimensions[total_row].height = 18
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
    for col in (1, 2):
        c = ws.cell(row=total_row, column=col)
        c.fill = mk_fill(dark_hex); c.border = thin_border()
    c = ws.cell(row=total_row, column=1, value="GRAND TOTAL")
    c.font = Font(name="Arial", bold=True, size=10, color=WHITE); c.alignment = CTR

    for col_idx in range(3, 3 + len(all_items)):
        col_letter = get_column_letter(col_idx)
        c = ws.cell(row=total_row, column=col_idx,
                    value=f"=SUM({col_letter}{data_start}:{col_letter}{data_end})")
        c.font = Font(name="Arial", bold=True, size=10, color=WHITE)
        c.fill = mk_fill(dark_hex); c.alignment = CTR; c.border = thin_border()
        c.number_format = '#,##0'

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 38
    for col_idx in range(3, 3 + len(all_items)):
        item = all_items[col_idx - 3]
        ws.column_dimensions[get_column_letter(col_idx)].width = max(10, min(18, len(item) * 0.85))
    ws.freeze_panes = "C4"

def build_location_tab(wb, sheet_title, location_name, items_dict, dark_hex, med_hex):
    ws = wb.create_sheet(title=sheet_title[:31])
    title_row(ws, 1, location_name, 3, dark_hex)
    info_row(ws, 2, f"Items shown: {len(items_dict)}   |   Negative qty never added into totals", 3, med_hex)

    ws.row_dimensions[3].height = 20
    header_cell(ws, 3, 1, "S.No", dark_hex, CTR)
    header_cell(ws, 3, 2, "Item Name", dark_hex, Alignment(horizontal="left", vertical="center"))
    header_cell(ws, 3, 3, "Balance Qty", dark_hex, CTR)

    sorted_items = sorted(items_dict.items(), key=lambda x: x[1], reverse=True)
    for row_num, (item, qty) in enumerate(sorted_items, start=1):
        er = 3 + row_num
        ws.row_dimensions[er].height = 16
        row_fill = mk_fill(ALT_GRAY) if row_num % 2 == 0 else mk_fill(WHITE)
        data_cell(ws, er, 1, row_num, row_fill, CTR)
        data_cell(ws, er, 2, item, row_fill, LFT)
        data_cell(ws, er, 3, qty, row_fill, CTR, '#,##0')

    n = len(sorted_items)
    total_r = 3 + n + 1
    ws.row_dimensions[total_r].height = 18
    for col, val in [(1, "TOTAL"), (2, "")]:
        c = ws.cell(row=total_r, column=col, value=val)
        c.font = Font(name="Arial", bold=True, size=10, color=WHITE)
        c.fill = mk_fill(dark_hex); c.border = thin_border(); c.alignment = CTR
    c = ws.cell(row=total_r, column=3, value=f"=SUM(C4:C{3+n})")
    c.font = Font(name="Arial", bold=True, size=10, color=WHITE)
    c.fill = mk_fill(dark_hex); c.border = thin_border(); c.alignment = CTR
    c.number_format = '#,##0'

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 16
    ws.freeze_panes = "A4"


def build_excel(data, out_path, title):
    bills = data['bills']
    locations = data['locations']
    parties = data['parties']
    all_items = data['all_items']

    # ── Aggregate by party — NEGATIVE QTY SKIPPED, NEVER ADDED ──
    party_data = defaultdict(lambda: defaultdict(float))
    for bill in bills:
        for item, qty in bill['items'].items():
            if qty > 0:
                party_data[bill['party']][item] += qty

    # ── Aggregate by location — same rule ──
    loc_data = defaultdict(lambda: defaultdict(float))
    for bill in bills:
        for item, qty in bill['items'].items():
            if qty > 0:
                loc_data[bill['location']][item] += qty

    def top_parties():
        totals = {p: sum(items.values()) for p, items in party_data.items() if sum(items.values()) > 0}
        return sorted(totals.items(), key=lambda x: x[1], reverse=True)[:20]

    def top_items():
        totals = defaultdict(float)
        for p, items in party_data.items():
            for item, qty in items.items():
                totals[item] += qty
        return sorted(totals.items(), key=lambda x: x[1], reverse=True)[:20]

    grand_total = sum(qty for p, items in party_data.items() for qty in items.values())

    wb = Workbook()
    wb.remove(wb.active)

    build_dashboard(
        wb, "Dashboard", f"📊 {title}",
        top_parties(), top_items(),
        len(parties), len(all_items), grand_total,
        DARK_BLUE, MED_BLUE, "2F75B5"
    )

    build_party_matrix(
        wb, "Party Wise", f"{title}  —  PARTY WISE",
        {p: dict(v) for p, v in party_data.items()},
        sorted(parties), all_items, DARK_BLUE, MED_BLUE
    )

    for loc in sorted(locations):
        items_dict = dict(loc_data[loc])
        build_location_tab(wb, loc, loc, items_dict, DARK_BLUE, MED_BLUE)

    wb.save(out_path)


# ════════════════════════════════════════════════════════
#  STREAMLIT UI
# ════════════════════════════════════════════════════════

st.title("📊 GSH Rental Balance Qty")
st.caption("ZIP upload karo (rental bill PDFs), Excel report bana ke milega. "
           "Negative qty kabhi sum me add nahi hoti.")

uploaded_zip = st.file_uploader("Rental bills ka ZIP file upload karo", type=["zip"])
report_title = st.text_input("Report title", value="GSH Rental Balance Qty")

if uploaded_zip is not None:
    if st.button("🚀 Excel Banao", type="primary"):
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "input.zip")
            with open(zip_path, "wb") as f:
                f.write(uploaded_zip.getbuffer())

            bills_dir = os.path.join(tmpdir, "bills")
            with zipfile.ZipFile(zip_path, "r") as z:
                z.extractall(bills_dir)

            progress_bar = st.progress(0)
            status_text = st.empty()

            def progress_cb(done, total, name):
                progress_bar.progress(done / total)
                status_text.text(f"Parsing {done}/{total}: {name}")

            with st.spinner("PDFs parse ho rahe hain..."):
                data = parse_folder(bills_dir, progress_cb)

            status_text.empty()
            progress_bar.empty()

            st.success(
                f"✅ Done — PDFs: {data['pdf_count']} | Locations: {len(data['locations'])} "
                f"| Parties: {len(data['parties'])} | Items: {len(data['all_items'])}"
            )
            if data['locations']:
                st.write("Locations: " + ", ".join(data['locations']))
            if data['no_balance_count']:
                st.info(f"{data['no_balance_count']} bills me Balance Qty nahi mila (final bills — skip kiye).")
            if data['errors']:
                st.warning(f"{len(data['errors'])} PDFs parse nahi hue.")

            out_path = os.path.join(tmpdir, "Rental_Balance_Qty.xlsx")
            build_excel(data, out_path, report_title)

            with open(out_path, "rb") as f:
                st.download_button(
                    "⬇️ Excel Download Karo",
                    data=f.read(),
                    file_name="Rental_Balance_Qty.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

st.divider()
st.caption("Data sirf is session me process hota hai, kahi save nahi hota. Har upload ek naya, alag run hai.")
