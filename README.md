# 📚 Grant Funding Data Pipeline

> This project automates the scraping, consolidation, and analysis of grant funding data from multiple Excel files. It produces a master sheet and a pivot table summarizing total education sector grant allocations by state.

---

## 📑 Table of Contents

* 📚 Overview
* 🏗️ Project Structure
* ⚙️ How to Run the Project
  * 🖥️ 1. Run the Notebook on Google Colab
  * 📂 2. Upload the Scraped Excel Files
  * 📝 3. Consolidate and Create the Pivot Table
* 📊 Analysis
* 🛠️ Technologies Used
* ✍️ Author


---

## 🏗️ Project Structure

- **grant-funding-pipeline.ipynb** — Main notebook for downloading Excel files using Playwright.

- **[Grant Funding Data Pipeline Folder](https://drive.google.com/drive/u/0/folders/1MnCNbtItDKzOJPJQcvJqJvAsIOfwNqGx)** — Folder location in google drive where README, documentation, and Excel files are saved.

- **[Grant Funding By State](https://docs.google.com/spreadsheets/d/1aQjTdy3WbWBGHwCO8_UDT0zoGhejCzfmHnyZ6g3qV8k/edit?gid=0#gid=0)** — Google Sheets file where data is imported, consolidated, and analyzed.

---

## ⚙️ How to Run the Project

### 🖥️ 1. Run the Notebook on Google Colab
- Open Google Colab.
- Upload or open the `grant-funding-pipeline.ipynb` notebook.
- Connect to a runtime (go to **Runtime > Connect**).
- Install necessary libraries (Playwright, OpenPyXL, Google APIs) inside the notebook if prompted.
- Run all cells (go to **Runtime > Run All**) or step through them.
- The notebook will scrape and download the Excel files automatically.

### 📂 2. Upload the Scraped Excel Files
- After running the notebook, the Excel files will appear in the `Grant Funding Data Pipeline Folder` inside Google Drive.
- The sheets are automatically uploaded to the spreadsheet `Grant Funding By State`.

### 📝 3. Consolidate and Create the Pivot Table
- Open the Google Apps Script editor (Extensions > Apps Script) inside your Google Sheet.
- Run the consolidateAllSheets() script.
- This script:
-- Merges all state tabs into one Master tab.
--Drops the first two rows from each imported sheet (headers).
--Adds a column labeling each row with its state.
- The pivot table summarizing total grant funding by state was created manually using Google Sheets' built-in pivot table tool, and is saved in the GFS Pivot tab.


## 📊 Analysis

After analyzing the data provided by the pipeline, it is easy to identify which states receive the most grant money.  
Additional insights into the specific purposes of these grants would provide even greater value for customers.

---

## 🛠️ Technologies Used

- **Python** (Google Colab)
- **Playwright** (for automated browser scraping)
- **Google Drive API** (for uploading/downloading files)
- **Google Sheets API** (for organizing and consolidating data)
- **Google Apps Script** (for Master tab creation)
- **openpyxl** (for working with Excel files)
- **asyncio** (for managing asynchronous operations)
- **OAuth2 Authentication** (for accessing Google services securely)

---

## ✍️ Author

Corey Jones
