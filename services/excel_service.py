import pandas as pd
import os
from datetime import datetime


def load_excel(folder, filename):
    df = pd.read_excel(os.path.join(folder, filename), header=2)

    df.columns = df.columns.str.strip()
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

    return df


def filter_by_days_in_stage(df, stage_name="4. S/O"):
    # 1. Replace the column value with the value before the first space
    stage_name = f"Days on stage {stage_name}"
    
    df = df.copy()  # avoid modifying original DataFrame
    df[stage_name] = (
        df[stage_name]
        .astype(str)
        .str.split(" ", n=1)
        .str[0]
        .astype("Int64")  # nullable integer
    )

    # 2. Filter rows where the value is <= 7
    filter = df[df[stage_name] <= 7]

    # 3. Return filtered DataFrame
    return filter


def filter_renewal_next_month(df, date_col="1-2. Effective Date (if 1-1 YES) *"):
    
    today = datetime.today()

    # Convert date column to datetime
    df[date_col] = pd.to_datetime(
        df[date_col],
        format="%d/%m/%Y",
        errors="coerce"
    )

    # Calculate next month start
    if today.month == 12:
        start_next_month = datetime(today.year + 1, 1, 1)
    else:
        start_next_month = datetime(today.year, today.month + 1, 1)

    # Calculate end of next month
    end_next_month = datetime(start_next_month.year, start_next_month.month + 1, 1)

    # Condition 1: Stage starts with "1-2"
    cond_stage_renewal = df["Stage"].astype(str).str.startswith("1-2")

    # Condition 2: Date is within next month
    cond_next_month = (
        (df[date_col] >= start_next_month) &
        (df[date_col] < end_next_month)
    )

    # Apply BOTH conditions
    filtered = df[
        cond_stage_renewal & cond_next_month
    ]

    return filtered


def detect_stage_changes(df_old, df_new):
    merged = df_new.merge(
        df_old[["ID", "Stage"]],
        on="ID",
        how="left",
        suffixes=("", "_old")
    )

    changed = merged[
        (merged["Stage"] != merged["Stage_old"]) &
        (~merged["Stage_old"].isna())
    ]

    new = merged[merged["Stage_old"].isna()]

    result = pd.concat([changed, new]).copy()
    
    renewal = filter_renewal_next_month(df_new)
    result = pd.concat([result, renewal]).copy()
    
    return result

def detect_owned_by_sales(df_new):
    # 1. Filter owned by Sales
    df_sales = df_new[
        df_new["Owner Fullname"].isin([
            "MGTI okada", 
            "MGTI Barri", 
            "MGTI Luky",
            "Katsuhiro Motono"
            ])
        ]

    # 2. Identify next-month renewals (Stage = "1-2" AND effective date next month)
    df_sales_renewal = filter_renewal_next_month(df_sales)
    df_sales_so = filter_by_days_in_stage(df_sales)
    df_sales_inv = filter_by_days_in_stage(df_sales, "5.  Sales (Invoice)")

    # 3. Remove renewals from df_sales
    df_sales_clean = df_sales[
        ~df_sales["Stage"].astype(str).str.startswith("1-2")
    ].copy()
    df_sales_clean = df_sales_clean[
        ~df_sales_clean["Stage"].astype(str).str.startswith("4.")
    ].copy()
    df_sales_clean = df_sales_clean[
        ~df_sales_clean["Stage"].astype(str).str.startswith("5.")
    ].copy()

    # 4. Return remaining sales records
    result = pd.concat([
        df_sales_clean, 
        df_sales_renewal, 
        df_sales_so,
        df_sales_inv
        ]).copy()
    return result.copy()
