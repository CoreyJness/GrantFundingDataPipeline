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

## 📚 Overview

This project automates the scraping, consolidation, and analysis of grant funding data from multiple Excel files.  
It produces a master sheet and a pivot table summarizing total education sector grant allocations by state.

---

## 🏗️ Project Structure

- **grant-funding-pipeline.ipynb** — Main notebook for downloading Excel files using Playwright.
- **Assessment Folder** — Folder location in google drive where README, documentation, and Excel files are saved.
- **GrantFundingSpreadsheet.xlsx** — Google Sheets file where data is imported, consolidated, and analyzed.

---

## ⚙️ How to Run the Project

### 🖥️ 1. Run the Notebook on Google Colab
- Open Google Colab.
- Upload or open the `Corey J Burbio Assessment.ipynb` notebook.
- Connect to a runtime (go to **Runtime > Connect**).
- Install necessary libraries (Playwright, OpenPyXL, Google APIs) inside the notebook if prompted.
- Run all cells (go to **Runtime > Run All**) or step through them.
- The notebook will scrape and download the Excel files automatically.

### 📂 2. Upload the Scraped Excel Files
- After running the notebook, the Excel files will appear in the `Corey J Assessment Folder` inside your Google Drive.
- The sheets are automatically uploaded to the `Burbio Assessment Corey J Spreadsheet`.

### 📝 3. Consolidate and Create the Pivot Table
- Open the Google Apps Script editor (**Extensions > Apps Script**) inside your Google Sheet.
- Run the `consolidateAllSheets()` script:

```javascript
function consolidateAllSheets() {
  // Your Apps Script code here
}
```
---

This script:
- Merges all state tabs into one Master tab.
- Drops the first two rows from each imported sheet (headers).
- Adds a column labeling each row with its state.

The pivot table summarizing total grant funding by state was created manually using Google Sheets' built-in pivot table tool.  
It is saved in the "Pivot" sheet.

---


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
