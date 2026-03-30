from report_utils import (
    setup_translator,
    get_excel_files,
    get_latest_files,
    load_excel,
    detect_stage_changes,
    build_report,
    save_report,
    filter_by_owner
)

def main():
    folder = "data"

    # === Translator ===
    translator = setup_translator()

    # === Get latest 2 files ===
    files = get_excel_files(folder)
    old_file, new_file = get_latest_files(files)

    print(f"Old file: {old_file}")
    print(f"New file: {new_file}")

    # === Load data ===
    df_old = load_excel(folder, old_file)
    df_new = load_excel(folder, new_file)

    # =========================================================
    # 1️⃣ CHANGE REPORT (based on comparison)
    # =========================================================
    changed = detect_stage_changes(df_old, df_new)

    print(f"Total changed/new deals: {len(changed)}")

    if not changed.empty:
        report_lines = build_report(changed, translator)
        save_report(report_lines, "プロジェクト_ステージ変更リスト.txt")
    else:
        print("⚠️ No changes detected")

    # =========================================================
    # 2️⃣ MGTI FULL REPORT (🔥 from latest file, NOT changes)
    # =========================================================
    mgt_df = filter_by_owner(df_new, "MGTI")

    print(f"MGTI total deals (all): {len(mgt_df)}")

    if not mgt_df.empty:
        mgt_report = build_report(mgt_df, translator)
        save_report(mgt_report, "プロジェクト_営業部管轄リスト.txt")
    else:
        print("⚠️ No MGTI deals found")

    print("✅ All reports generated successfully")


# === Entry point ===
if __name__ == "__main__":
    main()