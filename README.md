# QAD Automation Tool

This tool automates the process of logging into QAD Enterprise Applications, exporting data to Excel, and analyzing customer demand data with BOM information.

## Features

- Automated login to QAD using URL protocol
- Handling of protocol dialog with multiple methods
- Export data to Excel using keyboard navigation
- Analyze customer demand data with BOM information
- Generate component demand reports
- Robust error handling and logging
- Process verification with optional force mode

## Requirements

- Python 3.8+
- Edge WebDriver
- QAD Enterprise Applications
- Required Python packages:
  - selenium
  - pywinauto
  - keyboard
  - psutil
  - pandas
  - pyodbc
  - openpyxl

## Installation

1. Clone the repository
2. Install required packages:

```
pip install selenium pywinauto keyboard psutil pandas pyodbc openpyxl
```

## Usage

### Full QAD Automation

```
python run_full_automation.py --username <username> --password <password> [--state-id <state-id>] [--force]
```

#### Parameters

- `--username`: QAD username
- `--password`: QAD password
- `--state-id`: QAD state ID for custom folder navigation
- `--force`: Force execution even if QAD processes are running

### Data Analysis (Earlier Script)

```
python "Earlier Scripts/analyze_demand.py" [--excel-dir <directory>] [--sql-file <file>] [--output <file>] [--db-server <server>] [--db-name <database>] [--verbose]
```

#### Parameters

- `--excel-dir`: Directory containing the exported Excel files (default: Shell temp directory)
- `--sql-file`: SQL file with BOM queries (default: BOMs.sql)
- `--output`: Output file for component demand report (default: component_demand.xlsx)
- `--db-server`: Database server name
- `--db-name`: Database name (default: QADEE)
- `--verbose`: Enable verbose logging

### Environment Variables

You can also set credentials using environment variables:

- `QAD_USERNAME`: QAD username
- `QAD_PASSWORD`: QAD password

## Project Structure

- **run_full_automation.py**: Main script that integrates QAD login, data export, and analysis
- **Workflow.md**: Detailed workflow diagram of the automation process
- **URLs.md**: Contains QAD URLs for automation
- **Earlier Scripts/**: Contains previous versions of individual scripts
  - **qad-edge-automation.py**: Original script for QAD login and data export
  - **analyze_demand.py**: Original script for analyzing demand data
- **backup/**: Contains backup versions of scripts
  - **20250318_105714/**: Backups from March 18, 2025
    - **run_full_automation_backup_20250318_105714.py**
    - **Workflow_backup_20250318_105714.md**
  - **qad-edge-automation_backup_20250314_153017.py**: Backup from March 14, 2025
  - **Workflow_backup_20250314_153024.md**: Backup from March 14, 2025
- **logs/**: Contains log files from automation runs
- **sql querries/**: Contains SQL queries for data analysis

## Important Notes

- The script requires administrator privileges to interact with QAD windows.
- You need a valid QAD state ID to use with the script. This is the identifier for your specific QAD session.
- **During script execution, do not interact with the computer (mouse or keyboard).** The script simulates keyboard and mouse actions, and any manual interaction may interfere with the automation process.
- The script will automatically handle the protocol dialog that appears when opening QAD via URL.
- If QAD is already running, the script will detect it and can either continue (with `--force` option) or exit.
- The data analysis script can work with mock data if database connection fails.

## Workflow

See [Workflow.md](Workflow.md) for a detailed diagram of the automation and analysis process.

## Data Analysis Process

1. **Export Data**: The QAD automation script exports data to an Excel file.
2. **Find Latest Export**: The analysis script finds the latest exported Excel file.
3. **Read Demand Data**: Customer demand data is read from the Excel file.
4. **Retrieve BOM Data**: BOM data is retrieved from the database using SQL queries.
5. **Calculate Component Demand**: Component demand is calculated by joining demand data with BOM data.
6. **Generate Report**: A component demand report is generated and saved to Excel.

## Troubleshooting

- If the script fails with "Please close all QAD windows and try again", either close all QAD windows or use the `--force` option.
- If the protocol dialog handling fails, try running the script again.
- If database connection fails, the analysis script will use mock data.
- For detailed logs, check the console output or use the `--verbose` option.

## Backup Strategy

The project maintains backups in the following manner:
- Automated backups are created in the `backup/` folder with timestamps
- Major changes are committed to the Git repository
- Previous versions of individual scripts are preserved in the `Earlier Scripts/` folder

## URLs

QAD URLs are stored in [URLs.md](URLs.md) for reference.
