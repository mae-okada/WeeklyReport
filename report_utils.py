import pandas as pd
import os
import re

from datetime import datetime

# === Translator ===
def setup_translator():
    try:
        from deep_translator import GoogleTranslator
        return GoogleTranslator(source='auto', target='ja')
    except ImportError:
        print("⚠️ deep-translator not installed, using original text")
        return None


# === File handling ===
def get_excel_files(folder):
    pattern = re.compile(r"deals_\d{4}_\d{2}_\d{2}\.xlsx$")
    
    files = [
        f for f in os.listdir(folder)
        if pattern.match(f)
    ]

    if len(files) < 2:
        raise ValueError("❌ Need at least 2 valid deals_YYYY_MM_DD.xlsx files")

    return files

def extract_date(filename):
    match = re.search(r"deals_(\d{4})_(\d{2})_(\d{2})", filename)
    if match:
        return datetime.strptime(match.group(0), "deals_%Y_%m_%d")
    
    raise ValueError(f"Invalid filename format: {filename}")


def get_latest_files(files):
    files_sorted = sorted(files, key=extract_date)
    
    # Take last 2 (latest)
    return files_sorted[-2], files_sorted[-1]


# === Load ===
def load_excel(folder, filename):
    df = pd.read_excel(os.path.join(folder, filename), header=2)

    df.columns = df.columns.str.strip()
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

    return df


# === Detect changes ===
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
    ].copy()

    changed["Stage"] = changed["Stage"].astype(str).str.strip()

    print(f"Changed rows: {len(changed)}")
    return changed


# === Translation ===
def translate_project(text, translator):
    text_lower = text.lower()

    # if "support" in text_lower:
    #     return "ITサポート"
    # if "microsoft" in text_lower or "365" in text_lower:
    #     return "Microsoft 365導入"
    # if "esign" in text_lower or "sign" in text_lower:
    #     return "電子契約"

    if translator:
        try:
            return translator.translate(text)
        except:
            return text

    return text


# === Helpers ===
def to_juta(value):
    if pd.isna(value):
        return "-"
    return f"{int(value / 1_000_000)} Juta"


def format_row(row, translator):
    company_raw = row.get("Company", "-")
    company = clean_company_name(company_raw)
    size = to_juta(row.get("Size", 0))
    stage = str(row.get("Stage", ""))

    project_raw = str(row.get("Name", "-"))
    project_clean = clean_project_name(project_raw)
    project_jp = translate_project(project_clean, translator)

    text = f"・{company} ： {size} / {project_jp}"

    if not stage.startswith("1-"):
        text += " / <元野記入>"

    return text


# === Report builder ===
def build_report(changed_df, translator):
    stage_map = {
        "5.  Sales (Invoice)": "■ 注文書受領",
        "4. S/O": "■ 注文書受領",
        "3. Quotation": "■ 見積り提出",
        "2. Proposal": "■ 提案",
        "1-1. Potential (New Opportunity)": "■ 新規案件",
        "1-2. Potential (Renewal)": "■ 新規案件"
    }

    # 🔥 Group stages by JP title
    grouped = {}
    for stage, jp_title in stage_map.items():
        grouped.setdefault(jp_title, []).append(stage)

    report_lines = []

    # 🔥 Loop by JP title (NOT stage)
    for jp_title, stages in grouped.items():
        subset = changed_df[changed_df["Stage"].isin(stages)]

        if subset.empty:
            continue

        report_lines.append(jp_title)

        for _, row in subset.iterrows():
            report_lines.append(format_row(row, translator))

        report_lines.append("")

    return report_lines

# === Save ===
def save_report(lines, filename="output.txt"):
    with open(filename, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"✅ Report generated: {filename}")
    
def clean_company_name(name):
    if not isinstance(name, str):
        return "-"
    
    # Split at "-" and take first part
    cleaned = name.split("-")[0]
    
    # Trim whitespace
    return cleaned.strip()

def clean_project_name(name):
    if not isinstance(name, str):
        return "-"
    
    # Split only at FIRST " - "
    parts = name.split(" - ", 1)
    
    # If "-" exists → take the right side
    if len(parts) > 1:
        return parts[1].strip()
    
    return name.strip()