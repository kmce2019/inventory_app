# Asset Inventory and Reporting Tool

This is a graphical user interface (GUI) application for Windows designed to consolidate asset information from various spreadsheets and generate comprehensive inventory reports.

## Features

- **Consolidate Multiple Sources:** Upload a master employee list and multiple asset spreadsheets (computers, phones, etc.) to create a single source of truth.
- **Normalize User Data:** Matches various usernames (`jdoe`, `john.doe`) to a single employee `Position ID`.
- **Generate Multiple Reports:** Create and export different inventory views as `.csv` files with the click of a button.
- **Reports Include:**
  - Full inventory for all active employees.
  - Inventory of assets assigned to inactive/unmatched users.
  - Specific reports for computers, desk phones, and cell phones.
  - A report showing users with both a desk phone and a cell phone.
  - A report on desk phone usage duration for active employees.
- **Customizable Configuration:** Easily edit the script to change column names, asset keywords, and exclusion lists.

## Setup and Installation

1.  **Clone the repository:**
    ```bash
    git clone <your-repository-url>
    cd <your-repository-name>
    ```

2.  **Create a virtual environment (recommended):**
    ```bash
    python -m venv venv
    venv\Scripts\activate
    ```

3.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

## How to Use

1.  **Run the application:**
    ```bash
    python inventory_app.py
    ```
2.  **Upload Files:**
    - Click **"1. Upload Master Employee List"** to select your main employee spreadsheet.
    - Click **"2. Upload Asset Files"** to select one or more spreadsheets for computers, cell phones, etc.
    - Click **"3. Upload Phone Usage Report"** to select the summarized phone usage data.
3.  **Process Data:**
    - Click the green **"4. Load & Process Files"** button. The app will read all the data and prepare it. The report buttons below will become active.
4.  **Generate Reports:**
    - Click any of the report buttons to generate and save the corresponding `.csv` file. A "Save As" dialog will appear, allowing you to choose the location and name for the exported file.

## Configuration

Before running, you may need to adjust the configuration variables at the top of the `inventory_app.py` script to match your spreadsheet formats. This is the most important step for ensuring the app works correctly.

- `COMPUTER_KEYWORDS`, `DESK_PHONE_KEYWORDS`, etc.: Keywords used to identify asset types from the filenames of your asset sheets.
- `USER_COLUMN_IN_ASSETS`: The exact column header in your asset sheets that contains the employee's name.
- `ASSET_FIELD_MAPPING`: Maps the desired columns in the final report to the possible column names in your source files.
- `EXCLUDE_EXACT_USERS` & `EXCLUDE_KEYWORD_USERS`: Lists of usernames to filter out of the "Inactive User Inventory" report.
- `PHONE_USAGE_COLUMNS`: The exact column headers in your summarized phone usage report.
