# 📈 TikTok Affiliate Budget Tracker

A simple Excel-based budget tracker designed to help TikTok affiliates manage and track their daily, monthly, and yearly income and expenses. This project is built using Excel VBA to automate tasks like data entry, summary calculations, and visual color coding.

## ✨ Features

- **Daily Tracking:** A dedicated sheet to log daily affiliate income, other income, and expenses.
- **Automated Calculations:** Automatically calculates total income and net profit/loss for each entry.
- **Visual Color Coding:** Applies green for profit, red for loss, and white for break-even, making it easy to see your performance at a glance.
- **Monthly & Yearly Summaries:** Automatically compiles your daily data into monthly and yearly summaries.
- **Navigation:** User-friendly buttons to quickly switch between the daily, monthly, and yearly sheets.

## 🛠️ How to Use

1. **Download:** Clone or download this repository.
2. **Open:** Open the `Excel_VBA_Budget_Tracker.xlsm` file. You may need to enable macros if prompted.
3. **Initialize:** Click the **"Initialize"** button or run the `InitializeBudgetTracker` macro to set up the sheets and navigation buttons for the first time.
4. **Add Entries:** Use the **"Add Daily Entry"** button to log your income and expenses.
5. **View Summaries:** Use the navigation buttons to view your monthly and yearly summaries. Use the **"Calculate Monthly"** button to refresh the summary data.

## 📦 Project Structure

- `Excel_VBA_Budget_Tracker.xlsm`: The main Excel file containing the workbook and VBA macros.
- `src/Module1.bas`: The exported VBA code module containing all the project's logic.
- `.gitignore`: Specifies files to be ignored by Git (e.g., temporary Excel files).

## 👩‍💻 Contributing

Feel free to open an issue or submit a pull request if you have suggestions for improvements or new features.

---

### 3. `src/Module1.bas`
This folder contains the actual VBA code. You'll need to export your module from Excel's VBA editor.

To do this:
1. Open your Excel file and press `Alt + F11` to open the **VBA Editor**.
2. In the Project Explorer pane on the left, find your **Module1**.
3. Right-click on **Module1** and select **Export File...**.
4. Save the file as `Module1.bas` inside the `src` folder of your project.

The content of this file will be the code you provided, which is correctly formatted for a `.bas` file. You should also ensure the `Excel_VBA_Budget_Tracker.xlsm` file itself is committed to the repository so users can download and start using it immediately.
