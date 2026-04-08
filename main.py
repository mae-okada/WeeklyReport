import os
import pandas as pd

from datetime import datetime

from services.file_service import get_excel_files, get_latest_files
from services.excel_service import extract_one_stage, load_excel, detect_stage_changes, detect_owned_by_sales, drop_one_stage
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

    # # Change report
    # print("<<Start>> Update list")
    # changed = detect_stage_changes(df_old, df_new)
    # report = build_report(changed, translator, True, True)
    # print("<<Saving>> Update list")
    # save_report(report, f"{output_folder}/プロジェクト_ステージ変更リスト.txt")
    # print("<<End>> Update list")

    # # Owned by Sales report
    # print("<<Start>> Sales list")
    # owned_by_sales = detect_owned_by_sales(df_new)
    # sales_report = build_report(owned_by_sales, translator, True)
    # print("<<Saving>> Sales list")
    # save_report(sales_report, f"{output_folder}/【赤丸】プロジェクト_営業部管轄リスト.txt")
    # print("<<End>> Sales list")
    
    # Final - Owned by Sales Weekly report
    print("<<Start>> Weekly list")
    
    # Get data in stage "4. S/O" that is new (not included in the previous report)
    changed = detect_stage_changes(df_old, df_new)
    changed_so_stage = extract_one_stage(changed, "4. S/O")
    
    #remove stage "4. S/O" from the new data frame to avoid duplication
    dropped_so_stage = drop_one_stage(df_new, "4. S/O")
    
    # concat the changed stage "4. S/O" and the rest of the new data frame, then filter for owned by sales
    final_report = pd.concat([changed_so_stage, dropped_so_stage]).drop_duplicates().copy()
    owned_by_sales = detect_owned_by_sales(final_report)
    
    # Build report for projects owned by sales
    sales_report = build_report(owned_by_sales, translator, False, True)
    
    #save the report with today's date
    print("<<Saving>> Weekly list")
    today_str = datetime.today().strftime("%Y%m%d")
    save_report(sales_report, f"{output_folder}/週刊レポート_{today_str}.txt")
    print("<<End>> Weekly list")

if __name__ == "__main__":
    main()