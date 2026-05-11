# 📊 Excel Data Filtering & Export Automation Tool

A Python-based automation script that imports Excel data, filters specific records based on conditions, performs calculations, and exports the processed data into a clean CSV file.

---

## 🚀 Project Overview

This project is designed to automate Excel data processing workflows. It reads an Excel file, applies filters (e.g., specific campaign type like **EXACT**), calculates key performance metrics, and exports the cleaned dataset for further analysis.

---

## ⚙️ Features

- 📥 Import Excel (.xlsx) file using OpenPyXL  
- 🔍 Filter data based on conditions  
- 📊 Perform calculations (Spend, Clicks, Sales, ACOS, CPC)  
- 📤 Export processed data to CSV file  
- 🧠 Automated report generation  

---

## 🛠️ Technologies Used

- Python 🐍  
- OpenPyXL 📊  
- CSV module 📁  
- Excel (.xlsx)  

---

## 📌 Workflow

1. Load Excel file (`import_sheet_1.xlsx`)  
2. Read data from specific sheet  
3. Filter rows based on campaign type (EXACT)  
4. Calculate:
   - Total Spend  
   - Total Clicks  
   - Total Sales  
   - Average ACOS  
   - Average CPC  
5. Export cleaned data into `export.csv`  

---

## 📊 Output Example

| Date | Portfolio | Campaign | Targeting | Spend | Clicks | Sales |
|------|----------|----------|-----------|-------|--------|-------|
| xxxx | xxxx     | xxxx     | EXACT     | xx.xx | xx     | xx.xx |

---

## 📈 Business Logic

- **ACOS = Spend / Sales**
- **CPC = Spend / Clicks**
- **Final Bid = Dynamic formula based on ACOS & CPC**

---
