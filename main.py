import os

from services.file_service import get_excel_files, get_latest_files
from services.excel_service import load_excel, detect_stage_changes, detect_owned_by_sales
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
    print("<<Start>> Update list")
    changed = detect_stage_changes(df_old, df_new)
    report = build_report(changed, translator)
    print("<<Saving>> Update list")
    save_report(report, f"{output_folder}/プロジェクト_ステージ変更リスト.txt")
    print("<<End>> Update list")

    # Owned by Sales report
    print("<<Start>> Sales list")
    owned_by_sales = detect_owned_by_sales(df_new)
    sales_report = build_report(owned_by_sales, translator, True)
    print("<<Saving>> Sales list")
    save_report(sales_report, f"{output_folder}/プロジェクト_営業部管轄リスト.txt")
    print("<<End>> Sales list")

    print("Done!")


if __name__ == "__main__":
    main()