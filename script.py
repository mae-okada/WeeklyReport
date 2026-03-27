import pandas as pd
import os
import re

# === Optional translator (safe import) ===
try:
    from deep_translator import GoogleTranslator
    translator = GoogleTranslator(source='auto', target='ja')
    USE_TRANSLATOR = True
except ImportError:
    print("⚠️ deep-translator not installed, using original text")
    USE_TRANSLATOR = False

# === 1. Get Excel files ===
folder = "data"
files = [f for f in os.listdir(folder) if f.endswith(".xlsx")]

if len(files) < 2:
    raise ValueError("❌ Need at least 2 Excel files in /data folder")

# === 2. Extract date from filename ===
def extract_date(filename):
    match = re.search(r"deals_(\d{4})_(\d{2})_(\d{2})", filename)
    if match:
        y, m, d = match.groups()
        return f"{y}{m}{d}"
    raise ValueError(f"Invalid filename format: {filename}")

files_sorted = sorted(files, key=extract_date)

old_file = files_sorted[-2]
new_file = files_sorted[-1]

print(f"Old file: {old_file}")
print(f"New file: {new_file}")

# === 3. Load Excel (header row = row 3) ===
df_old = pd.read_excel(os.path.join(folder, old_file), header=2)
df_new = pd.read_excel(os.path.join(folder, new_file), header=2)

# === 3.1 Clean columns (IMPORTANT) ===
df_old.columns = df_old.columns.str.strip()
df_new.columns = df_new.columns.str.strip()

df_old = df_old.loc[:, ~df_old.columns.str.contains("^Unnamed")]
df_new = df_new.loc[:, ~df_new.columns.str.contains("^Unnamed")]

# Debug (optional)
# print("Columns:", df_new.columns.tolist())

# === 4. Merge ===
merged = df_new.merge(
    df_old[["ID", "Stage"]],
    on="ID",
    how="left",
    suffixes=("", "_old")
)

# === 5. Detect changed stage ===
changed = merged[
    (merged["Stage"] != merged["Stage_old"]) &
    (~merged["Stage_old"].isna())
].copy()

changed["Stage"] = changed["Stage"].astype(str).str.strip()

print(f"Changed rows: {len(changed)}")

# === 6. Stage mapping ===
stage_map = {
    "5.  Sales (Invoice)": "■ 注文書受領",
    "4. S/O": "■ 注文書受領",
    "3. Quotation": "■ 見積り提出",
    "2. Proposal": "■ 提案",
    "1-1. Potential (New Opportunity)": "■ 新規案件",
    "1-2. Potential (Renewal)": "■ 新規案件"
}

# === 7. Translation (HYBRID) ===
def translate_project(text):
    text_lower = text.lower()

    if "support" in text_lower:
        return "ITサポート"
    if "microsoft" in text_lower or "365" in text_lower:
        return "Microsoft 365導入"
    if "esign" in text_lower or "sign" in text_lower:
        return "電子契約"

    if USE_TRANSLATOR:
        try:
            return translator.translate(text)
        except:
            return text

    return text

# === 8. Helpers ===
def to_juta(value):
    if pd.isna(value):
        return "-"
    return f"{int(value / 1_000_000)} Juta"

def format_row(row):
    company = row.get("Company", "-")
    size = to_juta(row.get("Size", 0))

    # 🔥 NOW USING "Name" COLUMN
    project_name = str(row.get("Name", "-")).strip()

    stage = str(row.get("Stage", ""))

    project_jp = translate_project(project_name)

    text = f"・{company} ： {size} / {project_jp}"

    # Add suffix if NOT potential
    if not stage.startswith("1-"):
        text += " / <元野記入>"

    return text

# === 9. Build report ===
report_lines = []

for stage, jp_title in stage_map.items():
    subset = changed[changed["Stage"] == stage]

    if subset.empty:
        continue

    report_lines.append(jp_title)

    for _, row in subset.iterrows():
        report_lines.append(format_row(row))

    report_lines.append("")

# === 10. Save ===
with open("output.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(report_lines))

print("✅ Report generated: output.txt")