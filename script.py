from report_utils import (
    setup_translator,
    get_excel_files,
    get_latest_files,
    load_excel,
    detect_stage_changes,
    build_report,
    save_report
)

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