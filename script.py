import pandas as pd
import os
import re

# === Optional translator ===
def setup_translator():
    try:
        from deep_translator import GoogleTranslator
        return GoogleTranslator(source='auto', target='ja')
    except ImportError:
        print("⚠️ deep-translator not installed, using original text")
        return None


# === 1. File handling ===
def get_excel_files(folder="data"):
    files = [f for f in os.listdir(folder) if f.endswith(".xlsx")]
    if len(files) < 2:
        raise ValueError("❌ Need at least 2 Excel files in /data folder")
    return files


def extract_date(filename):
    match = re.search(r"deals_(\d{4})_(\d{2})_(\d{2})", filename)
    if match:
        y, m, d = match.groups()
        return f"{y}{m}{d}"
    raise ValueError(f"Invalid filename format: {filename}")


def get_latest_files(files):
    files_sorted = sorted(files, key=extract_date)
    return files_sorted[-2], files_sorted[-1]


# === 2. Load & clean ===
def load_excel(folder, filename):
    df = pd.read_excel(os.path.join(folder, filename), header=2)

    # Clean columns
    df.columns = df.columns.str.strip()
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

    return df


# === 3. Detect changes ===
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


# === 4. Translation (Hybrid) ===
def translate_project(text, translator):
    text_lower = text.lower()

    if "support" in text_lower:
        return "ITサポート"

    if translator:
        try:
            return translator.translate(text)
        except:
            return text

    return text


# === 5. Helpers ===
def to_juta(value):
    if pd.isna(value):
        return "-"
    return f"{int(value / 1_000_000)} Juta"


def format_row(row, translator):
    company = row.get("Company", "-")
    size = to_juta(row.get("Size", 0))
    project_name = str(row.get("Name", "-")).strip()
    stage = str(row.get("Stage", ""))

    project_jp = translate_project(project_name, translator)

    text = f"・{company} ： {size} / {project_jp}"

    if not stage.startswith("1-"):
        text += " / <元野記入>"

    return text


# === 6. Report builder ===
def build_report(changed_df, translator):
    stage_map = {
        "5.  Sales (Invoice)": "■ 注文書受領",
        "4. S/O": "■ 注文書受領",
        "3. Quotation": "■ 見積り提出",
        "2. Proposal": "■ 提案",
        "1-1. Potential (New Opportunity)": "■ 新規案件",
        "1-2. Potential (Renewal)": "■ 新規案件"
    }

    report_lines = []

    for stage, jp_title in stage_map.items():
        subset = changed_df[changed_df["Stage"] == stage]

        if subset.empty:
            continue

        report_lines.append(jp_title)

        for _, row in subset.iterrows():
            report_lines.append(format_row(row, translator))

        report_lines.append("")

    return report_lines


# === 7. Save ===
def save_report(lines, filename="output.txt"):
    with open(filename, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"✅ Report generated: {filename}")


# === 8. MAIN ===
def main():
    folder = "data"

    translator = setup_translator()

    files = get_excel_files(folder)
    old_file, new_file = get_latest_files(files)

    print(f"Old file: {old_file}")
    print(f"New file: {new_file}")

    df_old = load_excel(folder, old_file)
    df_new = load_excel(folder, new_file)

    changed = detect_stage_changes(df_old, df_new)

    report_lines = build_report(changed, translator)

    save_report(report_lines)


if __name__ == "__main__":
    main()