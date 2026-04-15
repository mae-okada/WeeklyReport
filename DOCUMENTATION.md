# Weekly Report Generator Documentation

## Overview
This project reads the latest Excel deal files from the `data/` folder, compares the latest two files, and generates output reports in `output/`.

## Architecture
The project now uses a class-based design with clear responsibilities:

- `main.py`
  - `WeeklyReportApp`
    - `__init__(self, data_folder="data", output_folder="output")`
      - initializes service objects and folder paths
    - `today_str`
      - returns today's date formatted as `YYYYMMDD`
    - `create_output_folder(self)`
      - creates the `output/` folder if it does not exist
    - `load_latest_data(self)`
      - loads the two newest Excel files and returns old/new DataFrames
    - `generate_change_report(self, df_old, df_new)`
      - identifies stage and size changes
      - combines change records and saves the stage change report
    - `generate_weekly_report(self, df_old, df_new)`
      - constructs the weekly report for sales-owned projects
      - saves the weekly output file
    - `run(self)`
      - runs the end-to-end report generation workflow

- `services/file_service.py`
  - `ExcelFileLocator`
    - `__init__(self, folder)`
      - sets the folder to scan for Excel files
    - `get_excel_files(self)`
      - returns filenames matching `deals_YYYY_MM_DD.xlsx`
      - validates that at least two files exist
    - `extract_date(filename)`
      - parses the date from the filename
    - `get_latest_files(self, files)`
      - sorts files by date and returns the two latest

- `services/excel_service.py`
  - `ExcelService`
    - `load_excel(self, folder, filename)`
      - reads an Excel file and cleans header columns
    - `filter_by_days_in_stage(self, df, stage_name)`
      - filters rows with a stage and recent days-on-stage value
    - `filter_renewal_this_and_next_month(self, df, date_col)`
      - returns renewal rows with an effective date this or next month
    - `detect_stage_changes(self, df_old, df_new)`
      - detects rows whose stage changed or are new
      - excludes stage transitions from invoice stage
      - includes relevant renewals
    - `extract_one_stage(self, df, stage_prefix)`
      - selects rows matching a stage prefix
    - `drop_one_stage(self, df, stage_prefix)`
      - removes rows matching a stage prefix
    - `detect_owned_by_sales(self, df_new)`
      - includes all `4. S/O` stage rows regardless of owner
      - filters owned-by-sales rows and renewal projects for other stages
    - `detect_change_in_size(self, df_old, df_new, size_col="Size")`
      - detects rows with changed deal size values
    - `sort_by_size(self, df, ascending=False, size_col="Size")`
      - sorts dataframe by deal size
      - `ascending=False` (default): largest deals first
      - `ascending=True`: smallest deals first

- `services/report_service.py`
  - `ReportService`
    - `__init__(self, translator=None)`
      - accepts a translator service instance
    - `build_report(self, df, use_name=False, use_stage=False)`
      - groups rows by Japanese stage labels
      - formats each row for output
    - `save_report(self, lines, path)`
      - writes the report lines to a UTF-8 text file

- `services/translator.py`
  - `TranslatorService`
    - `__init__(self)`
      - initializes the translator backend if available
    - `_setup_translator(self)`
      - tries to import and configure GoogleTranslator
    - `translate_project(self, text)`
      - applies hard-coded translation rules
      - falls back to the translator backend or raw text

- `utils/formatter.py`
  - `format_row(row, translator, use_name=False, use_stage=True)`
    - formats each DataFrame row into a report line
    - includes company, project, size, stage, and owner details
    - uses translator service for project name translation

## Usage
1. Place Excel files in `data/` using the filename format `deals_YYYY_MM_DD.xlsx`.
2. Run the script with:

```bash
python main.py
```

3. Generated reports are saved in `output/`.

## Output Files
- `<YYYYMMDD>_ステージ変更リスト.txt`
- `<YYYYMMDD>_週刊レポート.txt`

## Extension Points
- Add new report types by adding methods to `WeeklyReportApp`
- Extend translation rules in `services/translator.py`
- Add new formatting rules in `utils/formatter.py`

## Notes
- If `deep_translator` is not installed, translation falls back to the original project name.
- The code uses pandas for Excel processing and relies on a stable `Stage` and `ID` structure in input files.
