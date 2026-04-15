import os
import re
from datetime import datetime

class ExcelFileLocator:
    def __init__(self, folder):
        self.folder = folder
        self.pattern = re.compile(r"deals_\d{4}_\d{2}_\d{2}\.xlsx$")

    def get_excel_files(self):
        files = [f for f in os.listdir(self.folder) if self.pattern.match(f)]

        if len(files) < 2:
            raise ValueError("❌ Need at least 2 valid files")

        return files

    @staticmethod
    def extract_date(filename):
        match = re.search(r"deals_(\d{4})_(\d{2})_(\d{2})", filename)
        return datetime.strptime(match.group(0), "deals_%Y_%m_%d")

    def get_latest_files(self, files):
        files_sorted = sorted(files, key=self.extract_date)
        return files_sorted[-2], files_sorted[-1]