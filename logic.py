from filename import create_file_name
import pandas as pd
import os
import re
from collections import Counter


days_want = 21


def is_data_clean(df):
    if 'Product Name' in df.columns:
        print("hello i am inside product name if condition ")
        return True

    # Check if first row looks like header (contains 'Product Name' or similar)
    first_row = df.iloc[0].astype(str).str.lower()
    if first_row.str.contains('product').any() or first_row.str.contains('name').any():
        return False

    return False


def clean_source_data(source_path):

    out = os.path.join(os.path.dirname(source_path), "final.xlsx")

    df = pd.read_excel(source_path)

    # Check if data is already clean
    if is_data_clean(df):
        # If already clean, just return the original path
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

    # Save final.xlsx next to the source file
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


def run_transfer(source_path, d1_path, d2_path, d2_name, out_path):
    cleaned_path = clean_source_data(source_path)

    # ── Read files ──────────────────────────────────────────────────────
    d2 = pd.read_excel(d2_path)
    d1 = pd.read_excel(d1_path)   # sam stock
    source = pd.read_excel(cleaned_path)  # Use the cleaned file

    # ── Lookup stock quantities ─────────────────────────────────────────
    lookup(source, d1, 'samakhosi stock')
    lookup(source, d2, f'{d2_name} stock')

    # ── Median across historical columns ───────────────────────────────
    mean_calc = source.columns[1:-2]
    source['median'] = source[mean_calc].median(axis=1)

    below_median = int((source['samakhosi stock'] < source['median']).sum())

    # ── Transfer calculation ─────────────────────────────────────────────
    result = []
    for _, row in source.iterrows():
        ans = row['median']

        # ans = ans * days_want

        # ans = int(min(
        #     int(min(
        #         row['samakhosi stock'] * 0.45,
        #         abs(row['median'] - row[f'{d2_name} stock'])
        #     )),
        #     row['median'] / 2
        # ))
        ans = int(min(ans, row['samakhosi stock']*0.45))

        if row[f'{d2_name} stock'] < ans:
            ans -= row[f"{d2_name} stock"]
        else:
            ans = 0

        if row['samakhosi stock'] < ans:
            result.append(0)
        else:
            if (ans <= 2 and row[f'{d2_name} stock'] == 0) or (ans == 1):
                ans = 0
            result.append(ans)

    # ── Write output ────────────────────────────────────────────────────
    source[f'{d2_name} transfer'] = result

    # Force the output file to live in the same folder as final.xlsx
    # (i.e. next to source_path), regardless of any directory passed in
    # via out_path. Only the filename portion of out_path is honored.

    out_filename = create_file_name(d2_name)

    out_path = os.path.join(
        os.path.dirname(source_path),
        out_filename
    )

    source.to_excel(out_path, index=False)

    # ── Return stats for the UI ─────────────────────────────────────────
    non_zero = sum(1 for v in result if v > 0)
    return {
        'total':        len(source),
        'below_median': below_median,
        'to_transfer':  non_zero,
        'skipped':      len(result) - non_zero,
        'out_path':     out_path,
    }
