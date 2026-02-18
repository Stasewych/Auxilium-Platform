# Payment Reconciliation Tool (Just Solution)

## 📌 Overview
This is a specialized web application developed for **Just Solution** to automate legal and financial workflows. It facilitates the reconciliation of debtor payments against accounting reports and provides tools for filtering IBAN lists to exclude protected accounts (e.g., salary or social payments).

## 🚀 Key Features

### 1. Payment Reconciliation (Звірка платежів)
Automates the process of matching bank payment reports with debtor registries.
- **Multi-Bank Support**:
    - **Original Format (Auxilium)**
    - **Idea Bank**
    - **Taskombank**
- **Smart Matching**: Identifies payments using a composite key of `Borrower Name + Contract Number`.
- **Automatic Reporting**: Generates a new Excel file with a dedicated column for the selected month's payments.
- **Deep Clean Technology**: Ensures downloaded files are free from "Shared Formula" errors by stripping internal metadata.

### 2. IBAN Filtering (Фільтрація IBAN)
Helps legal teams avoid blocking prohibited accounts by filtering them out of seizure lists.
- **Exclusion Logic**: Compares a "Master List" of all accounts against a "Safe List" (salary/social projects).
- **Statistics**: Provides real-time counts of total, excluded, and remaining accounts.
- **Clean Output**: Generates a file ready for executive proceedings.

## 🛠️ Technology Stack
- **Frontend**: HTML5, CSS3, JavaScript (ES6 Modules)
- **Data Processing**: `ExcelJS` (Advanced manipulation), `SheetJS` (Legacy support)
- **Architecture**: Modular Component-Based Structure

## 📦 Installation & Setup
This is a client-side web application. No server installation is required.

1.  **Download** the project folder.
2.  **Open** `index.html` in any modern web browser (Chrome, Edge, Firefox).
3.  **Login** with the provided credentials:
    -   **Login**: `AuxiliumUser`
    -   **Password**: `Auxilium2026!`

## 📖 Usage Guide

### Payment Reconciliation
1.  Go to the **"Звірка платежів"** tab.
2.  **Step 0**: Select your Bank (Original, Idea, or Taskombank).
3.  **Step 1**: Upload the **Debtor Registry** (the file to update).
4.  **Step 2**: Upload the **Accounting Report** (the source of funds).
5.  **Step 3**: Select the **Month** for the report.
6.  Click **"Виконати обробку"**.
7.  Download the result.

### IBAN Filtering
1.  Go to the **"Фільтрація IBAN"** tab.
2.  **Step 1**: Upload the **Master IBAN List**.
3.  **Step 2**: Upload the **Exclusion List** (salary cards).
4.  Check the statistics panel.
5.  Click **"Виконати фільтрацію"**.
6.  Download the clean list.

## ⚠️ Troubleshooting
-   **Download Error**: If you see an error about "Shared Formulas", try processing the file again. The system now includes an auto-fix for this.
-   **Encoding Issues**: If names look like random symbols, change the **"Кодування"** dropdown (try `Windows-1251` or `UTF-8`).

---
*Developed for internal use by Just Solution.*
