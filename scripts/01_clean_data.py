"""
01_clean_data.py
-----------------
Cleans the classic scale-model distributor sales dataset (motorcycles,
classic cars, planes, ships, trains sold to retailers worldwide, 2003-2005).

Key issue found during inspection: pandas' default CSV parsing treats the
string "NA" as a null value. This dataset uses "NA" as the literal code for
the North America TERRITORY, so a naive pd.read_csv() silently wiped out
1,074 of 2,823 rows worth of territory data (38% of the dataset) -- exactly
the rows needed for region-wise analysis. Fixed by disabling default NA
parsing for this column.
"""

import pandas as pd
import numpy as np

RAW_PATH = "sales_project/data/sales_raw.csv"
CLEAN_PATH = "sales_project/data/sales_clean.csv"

COST_FACTOR = 0.70  # assumed cost = 70% of MSRP; see README for rationale


def load_raw(path: str) -> pd.DataFrame:
    df = pd.read_csv(path, encoding="latin1", keep_default_na=False, na_values=[""])
    print(f"Raw shape: {df.shape}")
    return df


def clean(df: pd.DataFrame) -> pd.DataFrame:
    before = len(df)
    df = df.drop_duplicates()

    # Drop columns with no analytical value for this project (PII / mostly-empty)
    drop_cols = ["ADDRESSLINE1", "ADDRESSLINE2", "PHONE", "CONTACTFIRSTNAME",
                 "CONTACTLASTNAME", "POSTALCODE"]
    df = df.drop(columns=[c for c in drop_cols if c in df.columns])

    # Standardize TERRITORY: "NA" = North America; relabel for clarity and
    # avoid any downstream library re-interpreting "NA" as null again
    df["Territory"] = df["TERRITORY"].replace({"NA": "North America"})
    df = df.drop(columns=["TERRITORY"])

    # Parse order date
    df["Order Date"] = pd.to_datetime(df["ORDERDATE"], errors="coerce")
    df = df.dropna(subset=["Order Date"])
    df["Order Month"] = df["Order Date"].dt.to_period("M").astype(str)

    # --- Derived profit fields ---
    # This dataset has no direct cost/profit column. PRICEEACH is the actual
    # unit sale price; MSRP is the list price. We estimate unit cost as 70%
    # of MSRP (a standard distributor markup assumption, documented here and
    # in the README rather than presented as an exact figure).
    df["Est. Unit Cost"] = df["MSRP"] * COST_FACTOR
    df["Est. Profit"] = df["SALES"] - (df["Est. Unit Cost"] * df["QUANTITYORDERED"])
    df["Est. Profit Margin"] = np.where(df["SALES"] != 0,
                                         df["Est. Profit"] / df["SALES"], 0)

    # Rename a few columns to Title Case for readability downstream
    df = df.rename(columns={
        "ORDERNUMBER": "Order Number", "QUANTITYORDERED": "Quantity Ordered",
        "PRICEEACH": "Price Each", "SALES": "Sales", "STATUS": "Status",
        "PRODUCTLINE": "Product Line", "MSRP": "MSRP", "PRODUCTCODE": "Product Code",
        "CUSTOMERNAME": "Customer Name", "CITY": "City", "STATE": "State",
        "COUNTRY": "Country", "DEALSIZE": "Deal Size", "QTR_ID": "Quarter",
        "MONTH_ID": "Month", "YEAR_ID": "Year",
    })

    after = len(df)
    print(f"Clean shape: {df.shape} (removed {before - after} rows)")
    return df


if __name__ == "__main__":
    raw = load_raw(RAW_PATH)
    clean_df = clean(raw)
    clean_df.to_csv(CLEAN_PATH, index=False)
    print(f"Saved clean data to {CLEAN_PATH}")
    print(f"Total sales: ${clean_df['Sales'].sum():,.0f}")
    print(f"Est. total profit: ${clean_df['Est. Profit'].sum():,.0f}")
