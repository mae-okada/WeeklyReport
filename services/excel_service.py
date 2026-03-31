import pandas as pd
import os

def load_excel(folder, filename):
    df = pd.read_excel(os.path.join(folder, filename), header=2)

    df.columns = df.columns.str.strip()
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

    return df

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

    return pd.concat([changed, new]).copy()
