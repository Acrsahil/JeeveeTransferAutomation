"""
logic.py — Core transfer calculation logic.
Edit this file to change how calculations are made.
The UI (app.py) calls run_transfer() and expects the returned dict.
"""

import pandas as pd
import os


def is_data_clean(df):
    """
    Check if the data is already clean (has 'Product Name' column and proper format).

    Parameters
    ----------
    df : pandas.DataFrame   DataFrame to check

    Returns
    -------
    bool   True if data is already clean, False otherwise
    """
    # Check if 'Product Name' column exists
    if 'Product Name' in df.columns:
        print("hello i am inside product name if condition ")
        return True

    # Check if first row looks like header (contains 'Product Name' or similar)
    first_row = df.iloc[0].astype(str).str.lower()
    if first_row.str.contains('product').any() or first_row.str.contains('name').any():
        return False

    return False


def clean_source_data(source_path):
    """
    Clean the raw sales/source Excel file and save as final.xlsx
    beside the original file. If data is already clean, returns the original path.

    Parameters
    ----------
    source_path : str   Full path to the raw source .xlsx file

    Returns
    -------
    str   Full path to the cleaned file (or original if already clean)
    """
    # Check if cleaned file already exists
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


def lookup(from_tab, dest_tab, name):
    """Add a stock column to from_tab by matching product names in dest_tab."""
    qty = []
    for product in from_tab['Product Name']:
        match = dest_tab[dest_tab['Display Name'] == product]
        if not match.empty:
            value = match.iloc[0]['Free To Use Quantity']
        else:
            value = 0
        qty.append(value)
    from_tab[name] = qty


def run_transfer(source_path, d1_path, d2_path, d2_name, out_path):
    """
    Run the full transfer calculation.

    Parameters
    ----------
    source_path : str   Path to raw source.xlsx (will be cleaned automatically if needed)
    d1_path     : str   Path to d1.xlsx  (Samakhosi stock)
    d2_path     : str   Path to d2.xlsx  (destination stock)
    d2_name     : str   Transfer location name (e.g. "Bhaktapur")
    out_path    : str   Where to save the output .xlsx

    Returns
    -------
    dict with keys:
        total       : int   Total products in source
        below_median: int   Products where samakhosi stock < median
        to_transfer : int   Products with a non-zero transfer qty
        skipped     : int   Products skipped (transfer = 0)
        out_path    : str   Resolved output file path
    """
    # ── Clean the source file only if needed ──────────────────────────────
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
        ans = int(min(
            int(min(
                row['samakhosi stock'] * 0.45,
                abs(row['median'] - row[f'{d2_name} stock'])
            )),
            row['median'] / 2
        ))
        if row['samakhosi stock'] < ans:
            result.append(0)
        else:
            if (ans <= 2 and row[f'{d2_name} stock'] == 0) or (ans == 1):
                ans = 0
            result.append(ans)

    # ── Write output ────────────────────────────────────────────────────
    source[f'{d2_name} transfer'] = result

    if not out_path.endswith('.xlsx'):
        out_path += '.xlsx'
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