"""
GSH Rental Balance Qty — Web App
=================================
Upload a ZIP of rental bill PDFs -> get a formatted Excel back.

Rules Rahul asked for:
1. NEGATIVE quantities are never added into any sum. When the same party/item
   appears across multiple bills, only positive quantities are added together;
   negative lines are simply skipped (not subtracted, not summed).
2. The ZIP can contain ZIPs inside ZIPs inside RAR files, nested arbitrarily
   deep (this is how Rahul's real files are structured — outer ZIP has a ZIP
   per location, and each of those contains a RAR with the actual PDFs).
   Everything is extracted recursively before parsing.
3. Every item (SKU) has a known unit weight (from Rahul's W.T Sheet). Each
   report shows Wt (per unit) and Total Wt = Qty x Wt next to every item.

Run locally:
    pip install -r requirements.txt
    streamlit run app.py

Deploy for free (shareable link):
    Push this folder to a GitHub repo (including packages.txt), then deploy
    on share.streamlit.io.
"""
import os
import re
import glob
import shutil
import subprocess
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
#  ITEM WEIGHT TABLE  (from Rahul's W.T Sheet — per-unit kg)
#  Used to compute Total Wt = Qty x Wt for every item.
#  Item names are matched in UPPERCASE, trimmed.
# ════════════════════════════════════════════════════════

WEIGHT_MAP = {
    "ACRO SPAN 2.5 MTR": 40.0,
    "BASE JACK (350MM+450MM)": 3.2,
    "BASE JACK 350MM": 3.2,
    "BASE JACK 600MM": 3.2,
    "BASE JACK(650MM)": 3.2,
    "BASEJACK-450 MM": 3.2,
    "CHALI 8FT.": 16.0,
    "CHALI(6 TO 8)FT": 16.0,
    "CHALI-10 FT": 27.0,
    "CHALI-6 F'T": 15.0,
    "CHANNEL -9 FT": 20.0,
    "CHANNEL 10 FT": 22.0,
    "CHANNEL 20FT": 44.0,
    "CHANNEL 7 FT": 16.0,
    "CHANNEL 8 FT.": 18.0,
    "CHANNEL 8 TO 10 FT": 20.0,
    "CLAMP-MOVING": 0.85,
    "SPAN INNER": 20.0,
    "INNER -3M PROP": 10.0,
    "INNER 2M PROP": 7.0,
    "INNER 2360 MM": 7.5,
    "INNER 4MTR": 13.0,
    "INNER2.5MTR": 8.0,
    "JOINTER": 0.65,
    "LEDGER - 1/2 MM": 2.0,
    "LEDGER 1.2 MTR": 4.0,
    "LEDGER 1.5 MTR": 5.0,
    "LEDGER 1.8 M": 6.0,
    "LEDGER 1M": 3.5,
    "LEDGER 2M": 6.5,
    "LEDGER 915MM": 3.0,
    "MS PIPE 6 MTR": 20.0,
    "OUTER 2": 10.0,
    "OUTER 2 MTR": 10.0,
    "OUTER 3": 15.0,
    "SHUTTEIRNG PLATE-3*1.8": 15.0,
    "SHUTTERIN-2*2": 18.0,
    "SHUTTERING PLATE - 3*2 (HEAVY)": 21.0,
    "SHUTTERING PLATE 2*1.5(WELDED)": 10.0,
    "SHUTTERING PLATE 3*14(WELDED)": 15.0,
    "SHUTTERING PLATE -4*1.5": 21.0,
    "SHUTTERING PLATE 3*15": 12.0,
    "SHUTTERING PLATE 3*2": 21.0,
    "SHUTTERING PLATE 3*2 (17 KGS)": 17.0,
    "SHUTTERING PLATE 3X0.15": 10.0,
    "SHUTTERING PLATE 3X1": 10.0,
    "SHUTTERING PLATE 3X1.5": 15.0,
    "SHUTTERING PLATE 3X9": 10.0,
    "SHUTTERING PLATE-2*1.": 10.0,
    "SHUTTERING PLATE-2*1.8": 11.0,
    "SHUTTERING PLATE-2*2": 12.0,
    "SHUTTERING PROP 2*2": 18.0,
    "SHUTTERING PROP 2*4": 23.0,
    "SHUTTERING PROP 3*3": 25.0,
    "SHUTTERING PROP-2*3": 21.0,
    "SHUTTERING PROP-3*4": 28.0,
    "SHUTTERING PROP-2.5*2": 20.0,
    "SIKANJA": 2.0,
    "SIKANJA -4 FT.": 2.0,
    "SIKANJA 2.5 FT.": 2.0,
    "SIKANJA 2FT": 2.0,
    "SIKANJA 3.5 FT.": 2.0,
    "STANDARD 0.5 WITH 1 CUPS": 2.0,
    "STANDARD 1 MTR WITH 2 CUPS": 4.5,
    "STANDARD 1.5M WITH 3 CUPS": 7.0,
    "STANDARD 2 MTR WITH 4 CUPS": 9.0,
    "STANDARD 2.5 METER WITH 5 CUPS": 11.2,
    "STANDARD 3MTR WITH 6 CUPS": 13.2,
    "STEEL FORM-1.2": 22.0,
    "STEEL FORM-1.5 M": 25.0,
    "STEEL PLATE 2*1(WELDEED)": 10.0,
    "STEEL PLATE 2*18(WELDEED)": 12.0,
    "STEEL PLATE 2*2(WELDEED)": 10.0,
    "STEEL PLATE 3*.9(WELDED)": 10.0,
    "STEEL PLATE 3*1 (5.5 KGS)": 5.5,
    "STEEL PLATE 3*1(WELDED)": 10.0,
    "STEEL PLATE 3*1.15(WALDED)": 15.0,
    "STEEL PLATE 3*1.5 LIGHT WEIGHT": 15.0,
    "STEEL PLATE 3*1.5 WELDED": 15.0,
    "STEEL PLATE 3*15(WELDED)": 10.0,
    "STEEL PLATE 3*2 LIGHT WEIGHT": 17.0,
    "STEEL PLATE 3*21(WELDED)": 18.0,
    "STEEL PLATES 3* 18": 12.0,
    "STEEL PLATES 3*1.5 (14.5 KGS)": 14.5,
    "STEEL PLATES 3*18 (12 KGS)": 12.0,
    "STEEL PLATES 3*18(WELDED)": 12.0,
    "STEEL PLATES 3*2 (13 KGS)": 13.0,
    "STEEL PLATES 3*2 (17KG) NEW WELDED": 17.0,
    "STEEL PLATES 3*2 (17KG)WELDED": 17.0,
    "STEEL PLATES 3*2 (21 KGS)": 211.0,
    "UJACK (350+450)MM": 3.2,
    "UJACK 350MM": 3.2,
    "UJACK 600MM": 3.2,
    "UJACK-400 MM": 3.2,
    "UJACK-450 MM": 3.2,
    "WALKWAY CHALI-1.2 MTR": 10.0,
}

def get_weight(item_name):
    """Look up per-unit weight for an item. Returns None if not in the table."""
    return WEIGHT_MAP.get(item_name.strip().upper())


# ════════════════════════════════════════════════════════
#  ARCHIVE EXTRACTION — handles ZIP-inside-ZIP-inside-RAR,
#  arbitrarily deep, exactly like Rahul's real files.
# ════════════════════════════════════════════════════════

RAR_TOOLS = [
    ["unar", "-q", "-f", "-o"],   # unar <outdir> <file>  (installed via packages.txt)
    ["unrar", "x", "-y"],          # unrar x -y <file> <outdir>/
    ["7z", "x", "-y"],             # 7z x -y -o<outdir> <file>
]

def _extract_rar(fpath, extract_dir):
    os.makedirs(extract_dir, exist_ok=True)
    # unar
    try:
        r = subprocess.run(["unar", "-q", "-f", "-o", extract_dir, fpath],
                            capture_output=True, timeout=600)
        if r.returncode == 0:
            return True
    except FileNotFoundError:
        pass
    # unrar
    try:
        r = subprocess.run(["unrar", "x", "-y", fpath, extract_dir + os.sep],
                            capture_output=True, timeout=600)
        if r.returncode == 0:
            return True
    except FileNotFoundError:
        pass
    # 7z
    try:
        r = subprocess.run(["7z", "x", "-y", f"-o{extract_dir}", fpath],
                            capture_output=True, timeout=600)
        if r.returncode == 0:
            return True
    except FileNotFoundError:
        pass
    return False

def _find_and_extract_archives(root):
    """One pass: find every .zip/.rar anywhere under root and extract it
    into a same-named '<name>_ext' folder, then delete the archive.
    Returns True if anything was extracted (caller should call again)."""
    found = False
    for dirpath, _dirnames, filenames in list(os.walk(root)):
        for fname in filenames:
            lower = fname.lower()
            if not (lower.endswith(".zip") or lower.endswith(".rar")):
                continue
            fpath = os.path.join(dirpath, fname)
            extract_dir = os.path.join(dirpath, fname.rsplit(".", 1)[0] + "_ext")
            ok = False
            if lower.endswith(".zip"):
                try:
                    with zipfile.ZipFile(fpath) as z:
                        z.extractall(extract_dir)
                    ok = True
                except Exception:
                    ok = False
            else:
                ok = _extract_rar(fpath, extract_dir)
            if ok:
                try:
                    os.remove(fpath)
                except OSError:
                    pass
                found = True
    return found

def extract_all_archives(root, max_passes=15):
    """Keep extracting nested archives until nothing is left to extract."""
    for _ in range(max_passes):
        if not _find_and_extract_archives(root):
            break


# ════════════════════════════════════════════════════════
#  PARSING
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
#  EXCEL BUILDING
#  - Negative quantities are SKIPPED when combining bills — never added.
#  - Every item shows Wt (per unit) and Total Wt (Qty x Wt).
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
                     n_parties, n_items, grand_total, grand_weight, dark_hex, med_hex, bar_color):
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
        value=f"Parties: {n_parties}   |   Item Types: {n_items}   |   Grand Total Qty: {int(grand_total):,}"
              f"   |   Grand Total Weight: {int(grand_weight):,} kg")
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
                        sorted_parties, all_items, dark_hex, med_hex,
                        party_locations=None):
    """party_locations: optional dict party -> sorted list of location names.
    When given, an extra LOCATION column is inserted right after PARTY NAME
    so it's clear which location(s) each party's bills came from."""
    ws = wb.create_sheet(title=sheet_title)
    show_loc = party_locations is not None
    name_col = 2
    loc_col = 3 if show_loc else None
    item_start_col = 4 if show_loc else 3
    total_cols = (item_start_col - 1) + len(all_items) + 1
    weight_col = item_start_col + len(all_items)

    title_row(ws, 1, title_text, total_cols, dark_hex)
    info_row(ws, 2,
        f"Parties: {len(sorted_parties)}   |   Items: {len(all_items)}   |   Negative qty never added into totals",
        total_cols, med_hex)

    ws.row_dimensions[3].height = 50
    header_cell(ws, 3, 1, "S.No", dark_hex, CTR)
    header_cell(ws, 3, name_col, "PARTY NAME", dark_hex, Alignment(horizontal="left", vertical="center", wrap_text=True))
    if show_loc:
        header_cell(ws, 3, loc_col, "LOCATION", dark_hex, Alignment(horizontal="left", vertical="center", wrap_text=True))
    for idx, item in enumerate(all_items, start=item_start_col):
        c = ws.cell(row=3, column=idx, value=item)
        c.font = Font(name="Arial", bold=True, size=9, color=WHITE)
        c.fill = mk_fill(med_hex); c.alignment = WRP; c.border = thin_border()
    header_cell(ws, 3, weight_col, "TOTAL WEIGHT (kg)", dark_hex,
                Alignment(horizontal="center", vertical="center", wrap_text=True))

    for row_num, party in enumerate(sorted_parties, start=1):
        er = 3 + row_num
        ws.row_dimensions[er].height = 16
        row_fill = mk_fill(ALT_GRAY) if row_num % 2 == 0 else mk_fill(WHITE)
        data_cell(ws, er, 1, row_num, row_fill, CTR)
        data_cell(ws, er, name_col, party, row_fill, LFT)
        if show_loc:
            locs = ", ".join(party_locations.get(party, []))
            data_cell(ws, er, loc_col, locs, row_fill, LFT)
        p_items = party_data.get(party, {})
        party_weight = 0.0
        for col_idx, item in enumerate(all_items, start=item_start_col):
            qty = p_items.get(item, None)
            data_cell(ws, er, col_idx, qty, row_fill, CTR, '#,##0' if qty else None)
            if qty:
                wt = get_weight(item)
                if wt is not None:
                    party_weight += qty * wt
        data_cell(ws, er, weight_col, int(round(party_weight)) if party_weight else None,
                   row_fill, CTR, '#,##0')

    total_row = 3 + len(sorted_parties) + 1
    data_start, data_end = 4, 3 + len(sorted_parties)
    ws.row_dimensions[total_row].height = 18
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=item_start_col - 1)
    for col in range(1, item_start_col):
        c = ws.cell(row=total_row, column=col)
        c.fill = mk_fill(dark_hex); c.border = thin_border()
    c = ws.cell(row=total_row, column=1, value="GRAND TOTAL")
    c.font = Font(name="Arial", bold=True, size=10, color=WHITE); c.alignment = CTR

    for col_idx in range(item_start_col, item_start_col + len(all_items) + 1):
        col_letter = get_column_letter(col_idx)
        c = ws.cell(row=total_row, column=col_idx,
                    value=f"=SUM({col_letter}{data_start}:{col_letter}{data_end})")
        c.font = Font(name="Arial", bold=True, size=10, color=WHITE)
        c.fill = mk_fill(dark_hex); c.alignment = CTR; c.border = thin_border()
        c.number_format = '#,##0'

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions[get_column_letter(name_col)].width = 38
    if show_loc:
        ws.column_dimensions[get_column_letter(loc_col)].width = 20
    for col_idx in range(item_start_col, item_start_col + len(all_items)):
        item = all_items[col_idx - item_start_col]
        ws.column_dimensions[get_column_letter(col_idx)].width = max(10, min(18, len(item) * 0.85))
    ws.column_dimensions[get_column_letter(weight_col)].width = 16
    ws.freeze_panes = f"{get_column_letter(item_start_col)}4"

def build_location_tab(wb, sheet_title, location_name, items_dict, dark_hex, med_hex):
    ws = wb.create_sheet(title=sheet_title[:31])
    title_row(ws, 1, location_name, 5, dark_hex)
    info_row(ws, 2, f"Items shown: {len(items_dict)}   |   Negative qty never added into totals", 5, med_hex)

    ws.row_dimensions[3].height = 20
    header_cell(ws, 3, 1, "S.No", dark_hex, CTR)
    header_cell(ws, 3, 2, "Item Name", dark_hex, Alignment(horizontal="left", vertical="center"))
    header_cell(ws, 3, 3, "Balance Qty", dark_hex, CTR)
    header_cell(ws, 3, 4, "Wt (per unit)", dark_hex, CTR)
    header_cell(ws, 3, 5, "Total Wt", dark_hex, CTR)

    sorted_items = sorted(items_dict.items(), key=lambda x: x[1], reverse=True)
    for row_num, (item, qty) in enumerate(sorted_items, start=1):
        er = 3 + row_num
        ws.row_dimensions[er].height = 16
        row_fill = mk_fill(ALT_GRAY) if row_num % 2 == 0 else mk_fill(WHITE)
        wt = get_weight(item)
        data_cell(ws, er, 1, row_num, row_fill, CTR)
        data_cell(ws, er, 2, item, row_fill, LFT)
        data_cell(ws, er, 3, qty, row_fill, CTR, '#,##0')
        data_cell(ws, er, 4, wt, row_fill, CTR, '#,##0.00' if wt else None)
        data_cell(ws, er, 5, round(qty * wt, 2) if wt is not None else None, row_fill, CTR, '#,##0')

    n = len(sorted_items)
    total_r = 3 + n + 1
    ws.row_dimensions[total_r].height = 18
    for col, val in [(1, "TOTAL"), (2, ""), (4, "")]:
        c = ws.cell(row=total_r, column=col, value=val)
        c.font = Font(name="Arial", bold=True, size=10, color=WHITE)
        c.fill = mk_fill(dark_hex); c.border = thin_border(); c.alignment = CTR
    c = ws.cell(row=total_r, column=3, value=f"=SUM(C4:C{3+n})")
    c.font = Font(name="Arial", bold=True, size=10, color=WHITE)
    c.fill = mk_fill(dark_hex); c.border = thin_border(); c.alignment = CTR
    c.number_format = '#,##0'
    cw = ws.cell(row=total_r, column=5, value=f"=SUM(E4:E{3+n})")
    cw.font = Font(name="Arial", bold=True, size=10, color=WHITE)
    cw.fill = mk_fill(dark_hex); cw.border = thin_border(); cw.alignment = CTR
    cw.number_format = '#,##0'

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 13
    ws.column_dimensions["E"].width = 14
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

    # ── Which location(s) each party has bills from ──
    party_locations = defaultdict(set)
    for bill in bills:
        party_locations[bill['party']].add(bill['location'])
    party_locations = {p: sorted(locs) for p, locs in party_locations.items()}

    # ── Aggregate by (location, party) — for the per-location Party Wise tabs ──
    loc_party_data = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    for bill in bills:
        for item, qty in bill['items'].items():
            if qty > 0:
                loc_party_data[bill['location']][bill['party']][item] += qty

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
    grand_weight = sum(
        qty * get_weight(item) for p, items in party_data.items()
        for item, qty in items.items() if get_weight(item) is not None
    )

    wb = Workbook()
    wb.remove(wb.active)

    build_dashboard(
        wb, "Dashboard", f"📊 {title}",
        top_parties(), top_items(),
        len(parties), len(all_items), grand_total, grand_weight,
        DARK_BLUE, MED_BLUE, "2F75B5"
    )

    build_party_matrix(
        wb, "Party Wise", f"{title}  —  PARTY WISE (All Locations)",
        {p: dict(v) for p, v in party_data.items()},
        sorted(parties), all_items, DARK_BLUE, MED_BLUE,
        party_locations=party_locations
    )

    for loc in sorted(locations):
        items_dict = dict(loc_data[loc])
        build_location_tab(wb, loc, loc, items_dict, DARK_BLUE, MED_BLUE)

    # ── Per-location Party Wise tabs — same matrix, filtered to one location ──
    for loc in sorted(locations):
        loc_parties = sorted(loc_party_data[loc].keys())
        loc_items = sorted({item for p in loc_parties for item in loc_party_data[loc][p].keys()})
        build_party_matrix(
            wb, f"PW - {loc}"[:31], f"{title}  —  PARTY WISE  ({loc})",
            {p: dict(v) for p, v in loc_party_data[loc].items()},
            loc_parties, loc_items, DARK_BLUE, MED_BLUE
        )

    wb.save(out_path)

    unmatched = sorted({item for item in all_items if get_weight(item) is None})
    return unmatched


# ════════════════════════════════════════════════════════
#  STREAMLIT UI
# ════════════════════════════════════════════════════════

st.title("📊 GSH Rental Balance Qty")
st.caption("ZIP upload karo (ZIP ke andar ZIP/RAR bhi chalega, kitni bhi layers deep). "
           "Excel me har item ka Qty, Weight aur Total Weight milega. "
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

            with st.spinner("Andar ke ZIP/RAR files khol raha hoon (jitni bhi layers hon)..."):
                extract_all_archives(bills_dir)

            progress_bar = st.progress(0)
            status_text = st.empty()

            def progress_cb(done, total, name):
                progress_bar.progress(done / total)
                status_text.text(f"Parsing {done}/{total}: {name}")

            with st.spinner("PDFs parse ho rahe hain..."):
                data = parse_folder(bills_dir, progress_cb)

            status_text.empty()
            progress_bar.empty()

            if data['pdf_count'] == 0:
                st.error("Koi PDF nahi mila. ZIP ke andar ka structure check karo.")
            else:
                st.success(
                    f"✅ Done — PDFs: {data['pdf_count']} | Locations: {len(data['locations'])} "
                    f"| Parties: {len(data['parties'])} | Items: {len(data['all_items'])}"
                )
                if data['locations']:
                    st.write("Locations: " + ", ".join(data['locations']))
                if data['no_balance_count']:
                    st.info(f"{data['no_balance_count']} bills me Balance Qty nahi mila (final bills — skip kiye).")
                if data['errors']:
                    with st.expander(f"⚠️ {len(data['errors'])} PDFs parse nahi hue — dekho"):
                        for e in data['errors'][:50]:
                            st.text(e)

                out_path = os.path.join(tmpdir, "Rental_Balance_Qty.xlsx")
                unmatched = build_excel(data, out_path, report_title)

                if unmatched:
                    with st.expander(f"⚖️ {len(unmatched)} items ki weight table me nahi mili (Total Wt blank rahega)"):
                        for item in unmatched:
                            st.text(item)

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
