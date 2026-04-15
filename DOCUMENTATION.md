# Weekly Report Generator Documentation

## Overview
This project reads the latest Excel deal files from the `data/` folder, compares the latest two files, and generates output reports in `output/`.

## Architecture
The project now uses a class-based design with clear responsibilities:

- `main.py`
  - `WeeklyReportApp`
    - `__init__(self, data_folder="data", output_folder="output")`
      - initializes the service components required to load data, build reports, and translate text.
      - holds configuration for input and output folders.
    - `today_str`
      - returns the current date in `YYYYMMDD` format for file naming.
    - `create_output_folder(self)`
      - creates the `output/` directory if it does not already exist.
      - ensures the application can write reports safely.
    - `load_latest_data(self)`
      - finds the two newest deal files by date.
      - loads them into pandas DataFrames and returns old/new data sets.
    - `generate_change_report(self, df_old, df_new)`
      - detects projects whose stage or size changed.
      - merges change records and sorts them by deal size.
      - formats and saves the stage-change report.
    - `generate_weekly_report(self, df_old, df_new)`
      - isolates new or changed `4. S/O` stage records without owner filtering.
      - combines them with other sales-owned rows and renewals.
      - sorts the final dataset by deal size and saves the weekly report.
    - `run(self)`
      - orchestrates the full execution flow from folder setup to report generation.

- `services/file_service.py`
  - `ExcelFileLocator`
    - `__init__(self, folder)`
      - stores the path to the folder containing Excel files.
    - `get_excel_files(self)`
      - lists files matching the `deals_YYYY_MM_DD.xlsx` naming pattern.
      - validates that at least two files are present for comparison.
    - `extract_date(filename)`
      - parses the date portion from a matching filename.
      - returns a Python datetime object for sorting.
    - `get_latest_files(self, files)`
      - sorts matched files by parsed date.
      - returns the second-latest and latest filenames.

- `services/excel_service.py`
  - `ExcelService`
    - `load_excel(self, folder, filename)`
      - reads an Excel file starting from row 3 (`header=2`).
      - strips whitespace from column names and removes unnamed columns.
    - `filter_by_days_in_stage(self, df, stage_name)`
      - looks for the stage-specific `Days on stage` column.
      - converts it to numeric values and returns rows where the stage matches and days are 7 or fewer.
    - `filter_renewal_this_and_next_month(self, df, date_col)`
      - converts the renewal effective date to datetime.
      - selects `1-2` renewal rows with an effective date in this month or next month.
      - includes diagnostic prints for matched counts.
    - `detect_stage_changes(self, df_old, df_new)`
      - merges new and old data by `ID` and compares stage values.
      - selects rows with stage changes and new records.
      - excludes transitions from invoice-related stages.
      - adds renewal rows that meet the current effective date window.
    - `extract_one_stage(self, df, stage_prefix)`
      - returns rows where `Stage` begins with the provided prefix.
      - used to isolate `4. S/O` records explicitly.
    - `drop_one_stage(self, df, stage_prefix)`
      - removes rows where `Stage` begins with the provided prefix.
      - used to avoid duplicate `4. S/O` records in the weekly report.
    - `detect_owned_by_sales(self, df_new)`
      - selects all `4. S/O` rows regardless of owner.
      - additionally selects rows owned by the sales team from other stages.
      - includes renewals and removes duplicate IDs.
    - `detect_change_in_size(self, df_old, df_new, size_col="Size")`
      - compares deal sizes in new and old records by `ID`.
      - returns rows with size changes and newly added deals.
    - `sort_by_size(self, df, ascending=False, size_col="Size")`
      - converts the size column to numeric values.
      - sorts the DataFrame by deal size.
      - default behavior is descending order to show largest deals first.

- `services/report_service.py`
  - `ReportService`
    - `__init__(self, translator=None)`
      - stores a translator instance used when formatting project names.
    - `build_report(self, df, use_name=False, use_stage=False)`
      - groups rows by Japanese stage labels from configuration.
      - formats each row using `utils.formatter.format_row`.
      - includes placeholder output for empty groups.
    - `save_report(self, lines, path)`
      - writes the report lines to a UTF-8 text file.
      - preserves line breaks for plain-text output.

- `services/translator.py`
  - `TranslatorService`
    - `__init__(self)`
      - initializes the translation backend if available.
    - `_setup_translator(self)`
      - attempts to import `deep_translator.GoogleTranslator`.
      - falls back to `None` if translation is unavailable.
    - `translate_project(self, text)`
      - applies special-case translations for known product keywords.
      - strips prefixes like `NDID - ` before translation.
      - uses the backend translator for other text when available.

- `utils/formatter.py`
  - `format_row(row, translator, use_name=False, use_stage=True)`
    - formats each DataFrame row into a report line.
    - cleans company and project names.
    - converts and formats size values, including size differences.
    - appends stage-specific annotations, renewal dates, and owner info.
    - optionally adds stage transition labels when `use_stage=True`.

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
