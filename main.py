import os

from services.file_service import get_excel_files, get_latest_files
from services.excel_service import load_excel, detect_stage_changes
from services.report_service import build_report, save_report
from services.translator import setup_translator


def main():
    data_folder = "data"
    output_folder = "output"
    os.makedirs(output_folder, exist_ok=True)

    translator = setup_translator()

    files = get_excel_files(data_folder)
    old_file, new_file = get_latest_files(files)

    df_old = load_excel(data_folder, old_file)
    df_new = load_excel(data_folder, new_file)

    # Change report
    changed = detect_stage_changes(df_old, df_new)
    if not changed.empty:
        report = build_report(changed, translator)
        save_report(report, f"{output_folder}/プロジェクト_ステージ変更リスト.txt")

    # Full report
    report_all = build_report(df_new, translator)
    save_report(report_all, f"{output_folder}/プロジェクト_営業部管轄リスト.tx")

    print("✅ Done")


if __name__ == "__main__":
    main()