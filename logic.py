"""
logic.py — Core transfer calculation logic.
Edit this file to change how calculations are made.
The UI (app.py) calls run_transfer() and expects the returned dict.
"""

import pandas as pd
import numpy as np
import os
from sklearn.linear_model import LinearRegression

DAYS_TO_PREDICT   = 15
MONTH_DAYS        = 30
TRANSFER_FRACTION = 0.45


def is_data_clean(df):
    if 'Product Name' in df.columns:
        print("hello i am inside product name if condition ")
        return True
    first_row = df.iloc[0].astype(str).str.lower()
    if first_row.str.contains('product').any() or first_row.str.contains('name').any():
        return False
    return False


def clean_source_data(source_path):
    out = os.path.join(os.path.dirname(source_path), "final.xlsx")
    df = pd.read_excel(source_path)
    if is_data_clean(df):
        return source_path
    df = df.drop(1, axis=0)
    df = df.drop(2, axis=0)
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)
    df = df.iloc[:, :-1]
    df.rename(columns={df.columns[0]: "Product Name"}, inplace=True)
    df['Product Name'] = df['Product Name'].str.strip()
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


def forecast_15d(row, monthly_cols):
    """Predict 15-day demand for a single product row using linear regression."""
    sales = np.where(
        np.isfinite(row[monthly_cols].values.astype(float)),
        row[monthly_cols].values.astype(float),
        0.0
    )
    months_num = np.arange(1, len(monthly_cols) + 1).reshape(-1, 1)
    model = LinearRegression().fit(months_num, sales)
    next_month = max(model.predict([[len(monthly_cols) + 1]])[0], 0)
    return (next_month / MONTH_DAYS) * DAYS_TO_PREDICT


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
        total           : int   Total products in source
        below_forecast  : int   Products where samakhosi stock < 15-day forecast
        to_transfer     : int   Products with a non-zero transfer qty
        skipped         : int   Products skipped (transfer = 0)
        out_path        : str   Resolved output file path
    """
    cleaned_path = clean_source_data(source_path)

    d2     = pd.read_excel(d2_path)
    d1     = pd.read_excel(d1_path)
    source = pd.read_excel(cleaned_path)

    lookup(source, d1, 'samakhosi stock')
    lookup(source, d2, f'{d2_name} stock')

    monthly_cols = source.columns[1:-2]

    below_forecast = int(0)

    result = []
    for _, row in source.iterrows():
        pred          = forecast_15d(row, monthly_cols)
        sam_stock     = row['samakhosi stock']
        dest_stock    = row[f'{d2_name} stock']

        transfer_limit = min(sam_stock * TRANSFER_FRACTION, pred / 2)
        ans = int(min(transfer_limit, max(0, pred - dest_stock)))

        if sam_stock < ans or (ans <= 2 and dest_stock == 0) or ans == 1:
            ans = 0

        result.append(ans)

    source[f'{d2_name} transfer'] = result

    if not out_path.endswith('.xlsx'):
        out_path += '.xlsx'
    source.to_excel(out_path, index=False)

    non_zero = sum(1 for v in result if v > 0)
    return {
        'total':          len(source),
        'below_forecast': below_forecast,
        'to_transfer':    non_zero,
        'skipped':        len(result) - non_zero,
        'out_path':       out_path,
    }