import pandas as pd
import os

def load_excel(folder, filename):
    df = pd.read_excel(os.path.join(folder, filename), header=2)

    df.columns = df.columns.str.strip()
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

    return df

import pandas as pd
from datetime import datetime

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

    # 3. Remove renewals from df_sales
    df_sales_exclude_renewal = df_sales[
        ~df_sales["Stage"].astype(str).str.startswith("1-2")
    ].copy()

    # 4. Return remaining sales records
    result = pd.concat([df_sales_exclude_renewal, df_sales_renewal]).copy()
    return result.copy()
