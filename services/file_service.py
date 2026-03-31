import os
import re
from datetime import datetime

def get_excel_files(folder):
    pattern = re.compile(r"deals_\d{4}_\d{2}_\d{2}\.xlsx$")
    files = [f for f in os.listdir(folder) if pattern.match(f)]

    if len(files) < 2:
        raise ValueError("❌ Need at least 2 valid files")

    return files

def extract_date(filename):
    match = re.search(r"deals_(\d{4})_(\d{2})_(\d{2})", filename)
    return datetime.strptime(match.group(0), "deals_%Y_%m_%d")

def get_latest_files(files):
    files_sorted = sorted(files, key=extract_date)
    return files_sorted[-2], files_sorted[-1]