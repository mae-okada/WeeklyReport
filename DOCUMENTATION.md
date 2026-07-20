# Weekly Report Generator Documentation

## Overview
This project reads the latest Excel deal files from the `data/` folder, compares the latest two files when needed, and generates plain-text reports in the `output/` folder.

## Setup
The project now includes a lightweight Windows setup flow:

- Run `install_dependencies.bat` to install the required Python packages.
- Run `generate_report.bat` to execute the report generator.
- Dependency definitions are stored in `requirements.txt` and `package.json`.

Required Python packages:
- `pandas`
- `openpyxl`
- `deep-translator`

## Architecture
The project uses a class-based design with clear responsibilities:

- `main.py`
  - `WeeklyReportApp`
    - `__init__(self, data_folder="data", output_folder="output")`
      - initializes the service components for file loading, Excel processing, translation, and report building.
    - `today_str`
      - returns the current date in `YYYYMMDD` format for output filenames.
    - `create_output_folder(self)`
      - creates the output folder if it does not already exist.
    - `load_latest_data(self)`
      - finds the latest Excel files and loads them into pandas DataFrames.
    - `generate_change_report(self, df_old, df_new)`
      - builds the stage-change report when enabled.
    - `generate_weekly_report(self, df_old, df_new)`
      - builds the weekly report used by the default workflow.
    - `run(self)`
      - runs the weekly report flow by default.

- `services/file_service.py`
  - `ExcelFileLocator`
    - lists Excel files in the data folder and selects the appropriate files for comparison.

- `services/excel_service.py`
  - `ExcelService`
    - reads Excel files and performs filtering for renewals, stage changes, and sales-owned rows.

- `services/report_service.py`
  - `ReportService`
    - assembles report lines and writes them as UTF-8 text files.

- `services/translator.py`
  - `TranslatorService`
    - applies keyword-based translations for known product names.
    - replaces the whole word `renewal` with `更新` before any fallback translation is attempted.

- `utils/formatter.py`
  - `format_row(row, translator, use_name=False, use_stage=True)`
    - formats each row into the final report text.

## Usage
1. Place Excel files in `data/` using a recognizable filename pattern.
2. Run `install_dependencies.bat` once to install the required packages.
3. Run `generate_report.bat` to create the reports.
4. Generated reports are saved in `output/`.

## Output Files
- `<YYYYMMDD>_ステージ変更リスト.txt`
- `<YYYYMMDD>_週刊レポート.txt`

## Notes
- The current default entrypoint generates the weekly report.
- The change-report generation path is still available in `main.py` and can be enabled by uncommenting the relevant call.
- The translator uses `deep_translator` when available, but falls back gracefully if it is not installed.
