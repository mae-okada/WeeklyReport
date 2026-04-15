import os
from datetime import datetime

import pandas as pd

from services.file_service import ExcelFileLocator
from services.excel_service import ExcelService
from services.report_service import ReportService
from services.translator import TranslatorService


class WeeklyReportApp:
    def __init__(self, data_folder="data", output_folder="output"):
        self.data_folder = data_folder
        self.output_folder = output_folder
        self.file_locator = ExcelFileLocator(data_folder)
        self.excel_service = ExcelService()
        self.translator_service = TranslatorService()
        self.report_service = ReportService(self.translator_service)

    @property
    def today_str(self):
        return datetime.today().strftime("%Y%m%d")

    def create_output_folder(self):
        os.makedirs(self.output_folder, exist_ok=True)

    def load_latest_data(self):
        files = self.file_locator.get_excel_files()
        old_file, new_file = self.file_locator.get_latest_files(files)

        df_old = self.excel_service.load_excel(self.data_folder, old_file)
        df_new = self.excel_service.load_excel(self.data_folder, new_file)
        return df_old, df_new

    def generate_change_report(self, df_old, df_new):
        print("<<Start>> Update list")
        changed_stage = self.excel_service.detect_stage_changes(df_old, df_new)
        changed_size = self.excel_service.detect_change_in_size(df_old, df_new)

        changed_size_full = df_new.merge(
            changed_size[["ID", "Size_old"]],
            on="ID",
            how="inner"
        )

        changed = pd.concat([changed_stage, changed_size_full]).drop_duplicates(subset=["ID"])

        lines = self.report_service.build_report(changed, use_name=True, use_stage=True)
        self.report_service.save_report(lines, f"{self.output_folder}/{self.today_str}_ステージ変更リスト.txt")
        print("<<End>> Update list")

    def generate_weekly_report(self, df_old, df_new):
        print("<<Start>> Weekly list")

        changed = self.excel_service.detect_stage_changes(df_old, df_new)
        changed_so_stage = self.excel_service.extract_one_stage(changed, "4. S/O")
        dropped_so_stage = self.excel_service.drop_one_stage(df_new, "4. S/O")

        final_report = pd.concat([changed_so_stage, dropped_so_stage]).drop_duplicates().copy()
        owned_by_sales = self.excel_service.detect_owned_by_sales(final_report)

        lines = self.report_service.build_report(owned_by_sales, use_name=False, use_stage=True)
        self.report_service.save_report(lines, f"{self.output_folder}/{self.today_str}_週刊レポート.txt")
        print("<<End>> Weekly list")

    def run(self):
        self.create_output_folder()
        df_old, df_new = self.load_latest_data()
        self.generate_change_report(df_old, df_new)
        self.generate_weekly_report(df_old, df_new)


if __name__ == "__main__":
    WeeklyReportApp().run()
