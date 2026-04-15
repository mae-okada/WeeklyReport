import pandas as pd
from datetime import datetime


class ExcelService:
    def load_excel(self, folder, filename):
        df = pd.read_excel(f"{folder}/{filename}", header=2)
        df.columns = df.columns.str.strip()
        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
        return df

    def filter_by_days_in_stage(self, df, stage_name):
        col = f"Days on stage {stage_name}"

        if col not in df.columns:
            return pd.DataFrame(columns=df.columns)

        df = df.copy()
        df[col] = (
            df[col]
            .astype(str)
            .str.split(" ", n=1)
            .str[0]
            .astype("Int64")
        )

        cond_stage = df["Stage"].astype(str).str.startswith(stage_name.split(".")[0])
        cond_days = df[col].notnull() & (df[col] <= 7)

        return df[cond_stage & cond_days].copy()

    def filter_renewal_this_and_next_month(self, df, date_col="1-2. Effective Date (if 1-1 YES) *"):
        df = df.copy()
        today = datetime.today()

        df[date_col] = pd.to_datetime(
            df[date_col],
            format="%d/%m/%Y",
            errors="coerce"
        )

        start_this_month = datetime(today.year, today.month, 1)
        if today.month == 12:
            end_next_month = datetime(today.year + 1, 2, 1)
        elif today.month == 11:
            end_next_month = datetime(today.year + 1, 1, 1)
        else:
            end_next_month = datetime(today.year, today.month + 2, 1)

        cond_stage_renewal = df["Stage"].astype(str).str.startswith("1-2")

        print(
            f"Filtering for renewals with effective date between "
            f"{start_this_month.date()} and {end_next_month.date()}"
        )
        cond_next_month = (
            (df[date_col] >= start_this_month) &
            (df[date_col] < end_next_month)
        )

        print(
            f"Total records: {len(df)}, Stage 1-2: {cond_stage_renewal.sum()}, "
            f"Next month: {cond_next_month.sum()}"
        )

        filtered = df[cond_stage_renewal & cond_next_month]
        print(f"Filtered renewals this/next month: {len(filtered)}")

        return filtered

    def detect_stage_changes(self, df_old, df_new):
        merged = df_new.merge(
            df_old[["ID", "Stage"]],
            on="ID",
            how="left",
            suffixes=("", "_old")
        )

        changed = merged[
            (merged["Stage"] != merged["Stage_old"]) &
            (~merged["Stage_old"].isna()) &
            (~merged["Stage_old"].astype(str).str.startswith("5"))
        ]

        new = merged[merged["Stage_old"].isna()]
        result = pd.concat([changed, new]).copy()

        renewal = self.filter_renewal_this_and_next_month(df_new)
        return pd.concat([result, renewal]).copy()

    def extract_one_stage(self, df, stage_prefix="4. S/O"):
        return df[df["Stage"].astype(str).str.startswith(stage_prefix)].copy()

    def drop_one_stage(self, df, stage_prefix="4. S/O"):
        return df[~df["Stage"].astype(str).str.startswith(stage_prefix)].copy()

    def detect_owned_by_sales(self, df_new):
        df_sales = df_new[df_new["Owner Fullname"].isin([
            "MGTI okada",
            "MGTI Barri",
            "MGTI Luky",
            "Katsuhiro Motono"
        ])]

        df_sales_renewal = self.filter_renewal_this_and_next_month(df_sales)

        df_sales_clean = df_sales[
            ~df_sales["Stage"].astype(str).str.startswith(("1-2"))
        ].copy()

        result = pd.concat([df_sales_clean, df_sales_renewal]).copy()
        return result.drop_duplicates(subset=["ID"]).copy()

    def detect_change_in_size(self, df_old, df_new, size_col="Size"):
        merged = df_new.merge(
            df_old[["ID", size_col]],
            on="ID",
            how="left",
            suffixes=("", "_old")
        )

        changed = merged[
            (merged[size_col] != merged[f"{size_col}_old"]) &
            (~merged[f"{size_col}_old"].isna())
        ]

        new = merged[merged[f"{size_col}_old"].isna()]
        return pd.concat([changed, new]).copy()

    def sort_by_size(self, df, ascending=False, size_col="Size"):
        """
        Sort dataframe by deal size.
        
        Args:
            df: DataFrame to sort
            ascending: If False (default), largest sizes first; if True, smallest first
            size_col: Column name containing the size values
        
        Returns:
            Sorted DataFrame
        """
        df = df.copy()
        df[size_col] = pd.to_numeric(df[size_col], errors="coerce")
        return df.sort_values(by=size_col, ascending=ascending, na_position="last").copy()
