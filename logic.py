from filename import create_file_name
import pandas as pd
import os
import re
import math
from copy import copy
from collections import Counter
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter


days_want = 25

# If (transfer + branch stock) is below this share of the max month qty,
# top up by TOPUP_SHARE of the shortfall.
LOW_THRESHOLD = 0.50
TOPUP_SHARE = 0.50

# Top-up can never push (transfer + branch stock) above this share of samakhosi stock
TOPUP_SAMAKHOSI_CAP = 0.45

# Share of the online median that gets transferred
# (only used for products with zero retail sales in the past months)
ONLINE_SHARE = 0.30


def is_data_clean(df):
    if 'Product Name' in df.columns:
        return True

    # Check if first row looks like header (contains 'Product Name' or similar)
    first_row = df.iloc[0].astype(str).str.lower()
    if first_row.str.contains('product').any() or first_row.str.contains('name').any():
        return False

    return False


def clean_source_data(source_path, out_name="final.xlsx"):
    """
    Cleans a sales export. The cleaned file is saved next to the source
    file under `out_name` (so retail and online don't overwrite each other).
    """
    out = os.path.join(os.path.dirname(source_path), out_name)

    df = pd.read_excel(source_path)

    # Check if data is already clean
    if is_data_clean(df):
        return source_path

    # Otherwise, clean the data
    df = df.drop(1, axis=0)
    df = df.drop(2, axis=0)

    # Make first row the header
    df.columns = df.iloc[0]

    # Remove that row
    df = df[1:].reset_index(drop=True)
    df = df.iloc[:, :-1]
    df.rename(columns={df.columns[0]: "Product Name"}, inplace=True)

    df['Product Name'] = df['Product Name'].str.strip()

    df.to_excel(out, index=False)
    return out


def normalize(s):
    """Lowercase, strip whitespace, collapse multiple internal spaces."""
    return re.sub(r'\s+', ' ', str(s).strip().lower())


def strip_code(s):
    """Remove a leading [CODE] prefix like '[000GA0414X7] '."""
    return re.sub(r'^\[.*?\]\s*', '', str(s)).strip()


def lookup(from_tab, dest_tab, name, verbose=True):
    """
    Matches from_tab['Product Name'] against dest_tab['Display Name']
    using three tiers, in order:
      1. exact string match
      2. normalized match (case/whitespace differences)
      3. code-stripped match (ignores the leading [CODE] on BOTH sides —
         catches cases where the same product has a different code in
         each sheet)

    Prints a breakdown of which tier each match came from, and for
    anything still unmatched, tries to tell you whether it's a
    CODE MISMATCH (name text exists in dest_tab under a different code)
    or TRULY MISSING (no similar text found at all).
    """
    qty = []
    unmatched = []
    match_method = []

    dest_names = dest_tab['Display Name'].astype(str)
    dest_norm = dest_names.apply(normalize)
    dest_nocode_norm = dest_names.apply(lambda x: normalize(strip_code(x)))

    for product in from_tab['Product Name']:
        p = str(product)
        p_norm = normalize(p)
        p_nocode_norm = normalize(strip_code(p))

        # Tier 1: exact
        match = dest_tab[dest_names == p]
        method = 'exact'

        # Tier 2: normalized (case/whitespace)
        if match.empty:
            match = dest_tab[dest_norm == p_norm]
            method = 'normalized'

        # Tier 3: code stripped from both sides (handles code mismatches)
        if match.empty:
            match = dest_tab[dest_nocode_norm == p_nocode_norm]
            method = 'code-stripped-both'

        if not match.empty:
            value = match.iloc[0]['Free To Use Quantity']
        else:
            value = 0
            unmatched.append(p)
            method = 'NONE'

        qty.append(value)
        match_method.append(method)

    from_tab[name] = qty

    if verbose:
        total = len(from_tab['Product Name'])
        print(f"[{name}] Total: {total} | Matched: {total - len(unmatched)} | Unmatched: {len(unmatched)}")
        print(f"[{name}] Match method breakdown:", dict(Counter(match_method)))

        if unmatched:
            print(f"[{name}] --- Diagnosing {len(unmatched)} unmatched (showing first 15) ---")
            for p in unmatched[:15]:
                core = normalize(strip_code(p))
                probe = core[:25] if len(core) >= 25 else core
                close = dest_nocode_norm[dest_nocode_norm.str.contains(re.escape(probe), na=False)]
                if len(close) > 0:
                    print(f"  CODE MISMATCH?  {p}\n    -> closest dest match: {dest_names[close.index[0]]}")
                else:
                    print(f"  TRULY MISSING?  {p}")

    return unmatched


def load_and_merge_sales(retail_path, online_path):
    """
    Cleans both sales files and merges them on product name, giving:
        Product Name | retail months... | online months...
    Products that exist only in the online file are kept (retail = 0).
    Returns (merged_df, retail_cols, online_cols).
    """
    retail = pd.read_excel(clean_source_data(retail_path, "final.xlsx"))
    online = pd.read_excel(clean_source_data(online_path, "final_online.xlsx"))

    r_months = [c for c in retail.columns if c != 'Product Name']
    o_months = [c for c in online.columns if c != 'Product Name']

    r_map = {c: f"{c} (retail)" for c in r_months}
    o_map = {c: f"{c} (online)" for c in o_months}
    retail = retail.rename(columns=r_map)
    online = online.rename(columns=o_map)
    r_cols, o_cols = list(r_map.values()), list(o_map.values())

    for c in r_cols:
        retail[c] = pd.to_numeric(retail[c], errors='coerce').fillna(0)
    for c in o_cols:
        online[c] = pd.to_numeric(online[c], errors='coerce').fillna(0)

    key = lambda x: normalize(strip_code(x))
    retail['_key'] = retail['Product Name'].apply(key)
    online['_key'] = online['Product Name'].apply(key)

    # one row per product in each file (sums duplicates)
    retail = retail.groupby('_key', as_index=False).agg(
        {'Product Name': 'first', **{c: 'sum' for c in r_cols}}
    )
    online = online.groupby('_key', as_index=False).agg(
        {'Product Name': 'first', **{c: 'sum' for c in o_cols}}
    ).rename(columns={'Product Name': '_online_name'})

    merged = retail.merge(online, on='_key', how='outer')
    merged['Product Name'] = merged['Product Name'].fillna(merged['_online_name'])
    merged[r_cols + o_cols] = merged[r_cols + o_cols].fillna(0)
    merged = merged[['Product Name'] + r_cols + o_cols].reset_index(drop=True)

    return merged, r_cols, o_cols


# ── Output colours (yellow = retail, green = online) ───────────────────────
RETAIL_FILL = "FFFF00"
ONLINE_FILL = "92D050"
SAM_STOCK_FILL = "5B9BD5"      # blue   → samakhosi stock
BRANCH_STOCK_FILL = "F4B183"   # orange → branch stock
STOCK_LABEL_FILL = "D9D9D9"    # grey   → the "STOCK" label cell


def format_output(out_path):
    """
    Pure formatting (no calculation changes). Call this AFTER anything that
    reads/re-saves the file with pandas (e.g. preview_output), because the
    extra label row on top would otherwise be read as the header
    ("Unnamed: 0", "Unnamed: 1", ...).

      - finds the '(retail)' / '(online)' month columns from the header row
      - inserts a new top row with merged + centered 'Retail' / 'Online'
      - strips the suffix from the month headers
      - colours retail columns yellow and online columns green
    """
    wb = load_workbook(out_path)
    ws = wb.active

    # Already formatted → do nothing
    if ws.merged_cells.ranges:
        wb.close()
        return

    retail_idx, online_idx = [], []
    for c in range(1, ws.max_column + 1):
        v = str(ws.cell(row=1, column=c).value or "")
        if re.search(r'\(retail\)\s*$', v, flags=re.I):
            retail_idx.append(c)
        elif re.search(r'\(online\)\s*$', v, flags=re.I):
            online_idx.append(c)

    # new row 1 → the real header moves to row 2
    ws.insert_rows(1)
    last_row = ws.max_row

    def fill(hex_color):
        return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

    for label, idx, color in (("RETAIL SALES", retail_idx, RETAIL_FILL),
                              ("ONLINE SALES", online_idx, ONLINE_FILL)):
        if not idx:
            continue
        first, last = idx[0], idx[-1]

        ws.merge_cells(start_row=1, start_column=first, end_row=1, end_column=last)
        cell = ws.cell(row=1, column=first)
        cell.value = label
        cell.fill = fill(color)
        cell.font = Font(bold=True, size=12)
        cell.alignment = Alignment(horizontal="center", vertical="center")

        for col_idx in idx:
            ws.cell(row=1, column=col_idx).fill = fill(color)
            h = ws.cell(row=2, column=col_idx)
            h.value = re.sub(r'\s*\((retail|online)\)\s*$', '', str(h.value), flags=re.I)
            for r in range(2, last_row + 1):
                ws.cell(row=r, column=col_idx).fill = fill(color)
            ws.column_dimensions[get_column_letter(col_idx)].width = 15

    # ── Stock columns: samakhosi = blue, branch = orange ───────────────
    stock_idx = []
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(row=2, column=c).value or "").strip().lower()
        if not h.endswith(" stock"):
            continue
        stock_idx.append(c)
        # "STOCK" is already written above, so the header is just the name
        raw = str(ws.cell(row=2, column=c).value).strip()
        name = re.sub(r'\s+stock\s*$', '', raw, flags=re.I)
        ws.cell(row=2, column=c).value = name[:1].upper() + name[1:]
        color = SAM_STOCK_FILL if h.startswith("samakhosi") else BRANCH_STOCK_FILL
        for r in range(2, last_row + 1):
            ws.cell(row=r, column=c).fill = fill(color)
        ws.column_dimensions[get_column_letter(c)].width = 16

    # Merged, centered "STOCK" label above the stock columns
    if stock_idx:
        first, last = stock_idx[0], stock_idx[-1]
        if last > first:
            ws.merge_cells(start_row=1, start_column=first, end_row=1, end_column=last)
        cell = ws.cell(row=1, column=first)
        cell.value = "STOCK"
        cell.font = Font(bold=True, size=12)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        for c in range(first, last + 1):
            ws.cell(row=1, column=c).fill = fill(STOCK_LABEL_FILL)

    # ── LAST STEP: sort rows by the transfer column, highest → lowest ───
    # Values AND cell styles move together, so every colour/highlight
    # (ours or anything added earlier) stays exactly as it is.
    t_col = None
    for c in range(1, ws.max_column + 1):
        if "transfer" in str(ws.cell(row=2, column=c).value or "").lower():
            t_col = c          # last matching column = "<location> Transfer"

    if t_col and last_row > 3:
        def _num(v):
            try:
                return float(v)
            except (TypeError, ValueError):
                return 0.0

        snap = []
        for r in range(3, last_row + 1):
            snap.append([(ws.cell(row=r, column=c).value, copy(ws.cell(row=r, column=c)._style))
                         for c in range(1, ws.max_column + 1)])

        snap.sort(key=lambda row: _num(row[t_col - 1][0]), reverse=True)  # stable

        for i, row in enumerate(snap, start=3):
            for c, (val, style) in enumerate(row, start=1):
                cell = ws.cell(row=i, column=c)
                cell.value = val
                cell._style = copy(style)

    wb.save(out_path)


def run_transfer(source_path, online_path, d1_path, d2_path, d2_name, out_path):
    # ── Read files ──────────────────────────────────────────────────────
    d2 = pd.read_excel(d2_path)
    d1 = pd.read_excel(d1_path)   # sam stock
    source, retail_cols, online_cols = load_and_merge_sales(source_path, online_path)

    # ── Flag freebies — kept in sheet but transfer qty forced to 0 later ─
    freebies_mask = source['Product Name'].astype(str).str.contains(
        'FREEBIES', case=False, na=False
    )
    print(f"[freebies] {freebies_mask.sum()} freebie row(s) found — transfer will be set to 0.")

    # ── Lookup stock quantities ─────────────────────────────────────────
    lookup(source, d1, 'samakhosi stock')
    lookup(source, d2, f'{d2_name} stock')

    # Negative stock at the destination is treated as 0
    dest_col = f'{d2_name} stock'
    source[dest_col] = pd.to_numeric(source[dest_col], errors='coerce').fillna(0).clip(lower=0)

    # ── Medians ─────────────────────────────────────────────────────────
    source['retail median'] = source[retail_cols].median(axis=1)
    source['online median'] = source[online_cols].median(axis=1)

    # Max selling month qty (used only for the retail top-up step below)
    source['max month qty'] = source[retail_cols].max(axis=1).fillna(0)

    # Online is used ONLY when retail sold nothing in all past months
    retail_total = source[retail_cols].sum(axis=1)
    online_total = source[online_cols].sum(axis=1)
    use_online = (retail_total <= 0) & (online_total > 0)
    print(f"[online] {int(use_online.sum())} product(s) have no retail sales — using online sales.")

    below_median = int((source['samakhosi stock'] < source['retail median']).sum())

    # ── RETAIL transfer calculation (original median logic, unchanged) ──
    result = []
    for _, row in source.iterrows():
        ans = row['retail median'] / 30

        ans = ans * days_want

        ans = int(min(ans, row['samakhosi stock'] * 0.45))

        if row[dest_col] < ans:
            ans -= row[dest_col]
        else:
            ans = 0

        if row['samakhosi stock'] < ans:
            result.append(0)
        else:
            if (ans <= 2 and row[dest_col] == 0) or (ans == 1):
                ans = 0
            result.append(ans)

    # ── Retail top-up step ──────────────────────────────────────────────
    # position = transfer + stock already at the branch
    # If position is below 50% of the max month qty, find the shortfall
    # (max month - position) and send an extra 50% of it,
    # but (transfer + branch stock) is always capped at 45% of samakhosi stock.
    result_topup = list(result)
    topped_up = 0
    for i, (_, row) in enumerate(source.iterrows()):
        max_qty = row['max month qty']
        if max_qty <= 0:
            continue

        position = result_topup[i] + row[dest_col]

        if position < max_qty * LOW_THRESHOLD:
            diff = max_qty - position
            extra = math.ceil(diff * TOPUP_SHARE)
            # Cap applies to transfer + stock already at the branch
            sam_cap = int(row['samakhosi stock'] * TOPUP_SAMAKHOSI_CAP)
            room = max(sam_cap - row[dest_col], 0)
            new_qty = min(result_topup[i] + extra, room)

            # Subtract the stock already at the branch so the column shows
            # only the qty that actually needs to be sent.
            # (Never lower than the original transfer.)
            final_qty = max(new_qty - row[dest_col], result[i])

            if final_qty > result_topup[i]:
                result_topup[i] = final_qty
                topped_up += 1

    print(f"[top-up] {topped_up} product(s) were below 50% of max month and got a top-up.")

    # Never transfer just 1 unit in the top-up column
    result_topup = [0 if v == 1 else v for v in result_topup]

    # ── ONLINE transfer: 30% of the online median (capped at 45% of samakhosi stock) ──
    # Only for products with zero retail sales. Overrides the retail result.
    basis = []
    for i, (_, row) in enumerate(source.iterrows()):
        if use_online.iloc[i]:
            target = math.ceil(row['online median'] * ONLINE_SHARE)
            # same 45% Samakhosi cap as retail: target never above 45% of samakhosi stock
            target = int(min(target, row['samakhosi stock'] * TOPUP_SAMAKHOSI_CAP))
            qty = int(max(target - row[dest_col], 0))     # subtract stock already at the branch
            result_topup[i] = qty
            basis.append('Online')
        else:
            basis.append('Retail')

    # ── Final rule: never transfer just 1 unit (retail or online) ───────
    result_topup = [0 if v == 1 else v for v in result_topup]

    # ── Write output ────────────────────────────────────────────────────
    # The top-up result is the final transfer quantity.
    transfer_col = f'{d2_name} Transfer'
    source[transfer_col] = result_topup
    source['basis'] = basis

    # Zero out freebies — they appear in the sheet with all stock numbers
    # intact but their transfer quantity is always forced to 0.
    source.loc[freebies_mask, transfer_col] = 0

    # ── Diagnostics (console only — helps find why transfers are 0) ─────
    _sam = pd.to_numeric(source['samakhosi stock'], errors='coerce').fillna(0)
    print("[diagnostic] total rows                         :", len(source))
    print("[diagnostic] retail median > 0                  :", int((source['retail median'] > 0).sum()))
    print("[diagnostic] samakhosi stock > 0                :", int((_sam > 0).sum()))
    print("[diagnostic] sells (median > 0) but sam stock 0 :", int(((source['retail median'] > 0) & (_sam <= 0)).sum()))
    print("[diagnostic] final transfer > 0                 :", int((source[transfer_col] > 0).sum()))

    # Output file lives in the same folder as the retail source file.
    # Only the generated filename is used.
    out_filename = create_file_name(d2_name)

    out_path = os.path.join(
        os.path.dirname(source_path),
        out_filename
    )

    # 'max month qty' and the two median columns are only used internally,
    # so they are dropped from the file.
    out_df = source.drop(columns=['max month qty', 'retail median', 'online median'])

    # Sort by the transfer column, highest first (stable, so ties keep order)
    out_df[transfer_col] = pd.to_numeric(out_df[transfer_col], errors='coerce').fillna(0).astype(int)
    out_df = out_df.sort_values(transfer_col, ascending=False, kind='stable').reset_index(drop=True)

    out_df.to_excel(out_path, index=False)

    # NOTE: colours + the merged "Retail" / "Online" row are applied later by
    # format_output(out_path), called from app.py after preview_output.

    # ── Return stats for the UI ─────────────────────────────────────────
    non_zero = sum(1 for v in source[transfer_col] if v > 0)
    return {
        'total':        len(source),
        'below_median': below_median,
        'to_transfer':  non_zero,
        'skipped':      len(source) - non_zero,
        'online_based': int(use_online.sum()),
        'out_path':     out_path,
    }
