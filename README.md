# TikTok Affiliate Budget Tracker

This repository contains the VBA (Visual Basic for Applications) source code for an Excel-based budget tracker. The tool is designed to help TikTok affiliates manage and track their daily, monthly, and yearly income and expenses with automated calculations and visual aids.

Since this repository only contains the raw VBA code, you will need to import it into your own Excel workbook to use the functionality.

---

### ✨ Features

* **Daily Tracking:** Code to create a sheet for logging daily affiliate income, other income, and expenses.
* **Automated Calculations:** Macros to automatically calculate total income and net profit/loss for each entry.
* **Visual Color Coding:** Applies green for profit, red for loss, and white for break-even, making it easy to see your performance at a glance.
* **Monthly & Yearly Summaries:** Macros to compile your daily data into monthly and yearly summary sheets.
* **User-Friendly Navigation:** Code to create buttons that allow you to quickly switch between the daily, monthly, and yearly sheets.

---

### 🛠️ How to Use

1.  **Create an Excel File:** Open a new, blank Excel workbook and save it as a Macro-Enabled Workbook (`.xlsm`).
2.  **Open the VBA Editor:** Press `Alt + F11` to open the VBA editor.
3.  **Import the Module:**
    * In the VBA editor, right-click on your workbook in the Project Explorer pane (e.g., "VBAProject (YourFileName.xlsm)").
    * Select **Import File...**
    * Navigate to the location where you've saved `Module1.bas` and select it.
4.  **Run the Setup Macro:**
    * After importing, you should see `Module1` appear under your workbook.
    * Open the Immediate Window by pressing `Ctrl + G`.
    * Type `InitializeBudgetTracker` and press `Enter`. This will create the necessary worksheets and buttons in your Excel file.
5.  **Start Tracking:** Close the VBA editor and use the newly created buttons to add daily entries and view your summaries.

---

### 📁 Repository Contents

* `src/Module1.bas`: The main VBA code module containing all the macros and functions for the budget tracker.
