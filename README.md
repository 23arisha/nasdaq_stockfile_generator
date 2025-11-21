# Stock Table Excel Automation

**Automate daily stock table extraction and Excel report generation**

## Overview

This project automates the manual process of collecting stock market data from NASDAQ tables (percent gainers, percent decliners, volume actives, net decliners) and generates a clean, formatted Excel report. Previously, this process required manually visiting the website, expanding hidden rows, and copying data—taking up to 2 hours per day.

With this tool, all tables are scraped, formatted, and exported to Excel in **seconds**.

## Features

* Scrapes multiple stock tables from NASDAQ website.
* Automatically clicks “Load More” buttons to reveal all data.
* Generates Excel files with:

  * Proper headers
  * Borders and alignment for readability
  * Date stamps for daily tracking
* Runs via a single command in CLI.

## Tech Stack

* **Python 3**
* **Selenium** – Web automation
* **OpenPyXL** – Excel creation and formatting
* **chromedriver_autoinstaller** – Auto-install ChromeDriver
* **Datetime, OS, Time** – Utilities

## Installation

1. Clone the repository:

   ```bash
   git clone git clone https://github.com/23arisha/nasdaq_stockfile_generator.git
   cd nasdaq_stockfile_generator
   ```
2. Install dependencies:

   ```bash
   pip install selenium openpyxl chromedriver_autoinstaller
   ```

## Usage

1. Run the script:

   ```bash
   python nasdata.py
   ```
2. The script will:

   * Open NASDAQ stock page
   * Click “Load More” buttons for all tables
   * Extract data from the tables
   * Generate an Excel file in `~/Downloads` with the current date

## Output

* Excel file named `52weekDDMMM.xlsx` (e.g., `52week21Nov.xlsx`)
* Contains 4 sheets:

  1. Percent Gainers
  2. Percent Decliners
  3. Net Decliners
  4. Volume Actives

