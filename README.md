# NASDAQ Stock Market Analytics & Automation Portfolio

## 👤 Author
**Shahrzad Khorrami**  
Excel SME | Data Automation Specialist  
📧 shahrzad.khorrami@gmail.com  
🌐 LinkedIn: https://www.linkedin.com/in/shahrzad-khorrami

---

## 📌 Project Goal

This project demonstrates an **end-to-end stock market analytics and automation solution** using real-world NASDAQ historical data.
https://www.kaggle.com/datasets/jacksoncrow/stock-market-dataset?resource=download

The main objectives are:

- Consolidate and clean **thousands of ticker-level CSV files** (stocks + ETFs) using Excel Power Query.
- Build an **optimized data model** with Power Pivot for advanced analysis.
- Apply **Excel advanced formulas** (SUMIFS, IFS, XLOOKUP, nested IFs, upper/lower, data validation, sorting).
- Automate workflow using **VBA** and **Power Automate** for data import, refresh, and reporting.
- Perform **advanced analytics** using MySQL.
- Build **interactive dashboards** in Excel and Power BI for business insights.

---

## 🗂 Dataset

- **Source:** Historical daily prices for all NASDAQ tickers (up to April 1, 2020) via Yahoo Finance (Python yfinance package).  
- **Structure:**  
  - CSV files per ticker under `stocks/` and `etfs/` folders  
  - Columns: `Date, Open, High, Low, Close, Adj Close, Volume`  
  - Metadata: `symbols_valid_meta.csv` (Symbol, Security Name, ...)

- **Business Use Case:** Transform raw stock market data into actionable insights and automate reporting for investment analysis.

---

## 🏗 Architecture

**Workflow:**

1. **Data Import & Cleaning:** Excel Power Query imports CSVs → appends stocks + ETFs → merges metadata → transforms & cleans data.
2. **Data Model:** Load cleaned table into Power Pivot → define relationships → calculate KPIs using DAX.
3. **SQL Analytics:** Advanced queries on MySQL for cumulative volume, rolling averages, ranking, sector performance.
4. **Excel Dashboard:** Interactive dashboards with slicers, charts, conditional formatting, pivot tables.
5. **VBA Automation:** Auto-refresh data, import new files, export dashboards to PDF, trigger events.
6. **Power BI Visualization:** Dynamic dashboards with KPIs, combo charts, heatmaps, waterfall charts, drillthrough and slicers.
7. **Power Automate:** Automatically refresh dashboards when new data is added and send email reports.

---

## 📊 Key Features

**Excel & Power Query**
- Import thousands of CSV files and append into a single data table.
- Data cleaning: remove blanks, errors, normalize tickers (upper/lower), merge metadata.
- Advanced formulas: `SUMIFS`, `IFS`, `XLOOKUP`, Filter, nested IFs, `UPPER/LOWER/PROPER`, Unique, `SEARCH`.
- Pivot tables, conditional formatting, data validation, sorting, and filtering.
- Calculated KPIs: Daily Return, Rolling 30-Day Avg, Cumulative Volume, Volatility ranking.

**VBA Automation**
Created automated macro system:

✔ Button: "Export to CSV and import to MySQL Database"
Exports ETFs Table to CSV
Excludes calculated columns
Macro imports CSV into MySQL
Automated execution
No manual SQL needed
Row Duplicate handling included

✔ Event-driven logic
Clears results if ticker input is empty
Refresh logic triggered automatically

This simulates real financial data ingestion pipelines.

**SQL (MySQL)**
Database: etf_portfolio
Table: etfs

Automated CSV export
Automated Imported into MySQL table via macro
Proper column mapping
Data types aligned

Handled:
ODBC 64-bit driver configuration
Connector/NET requirement
Driver troubleshooting

**Power BI Dashboards**
Imported ETF CSV into Power BI.

Steps:
Get Data → CSV
Data cleaning in Power Query (Power BI)
Built DAX measures
Created visual dashboards

Visuals Created:

✔ KPI Cards
✔ Line charts (Price trends)
✔ Column charts (Volume analysis)
✔ Slicers (Ticker, Date)
✔ Interactive filtering

Why Power BI?
Professional visualization layer
Advanced DAX calculations
Better executive-level presentation
Real BI environment simulation

**Power Automate**  
(TODO)
- Trigger: New CSV or database update.
- Refresh Excel + Power Query + Power Pivot.
- Export dashboards automatically.
- Send automated email reports.

---

## 📈 Business Impact

- Eliminated manual processing of thousands of stock files.
- Optimized large dataset handling via Power Query and Data Model.
- Automated reporting workflow → reduced errors and saved time.
- Enabled interactive dashboards for executive decision-making.
- Demonstrated integration of **Excel, VBA, SQL, Power BI, and Power Automate** for enterprise-level analytics.

---

## 📁 Repository Structure

nasdaq-stock-excel-sme-portfolio/

├── data/ # raw CSV files

├── excel/Stock_Dashboard.xlsm

├── sql/etf_portfolio.sql

├── powerbi/ETF_Dashboard_POWERBI.pbix, ETF_Dashboard_POWERBI.pdf

├── screenshots/ # dashboard, Power Query, VBA, Power BI

└── README.md

---

## 📸 Screenshots

Include screenshots in `/screenshots/` for:

- Power Query transformations
- Data Model relationships
- Excel Dashboard charts
- VBA editor + macros
- Power BI dashboards
- Power Automate workflow

---

## 💡 Notes

- The repository contains **sample CSV files** to keep size manageable; you can use full NASDAQ dataset locally.
- The project demonstrates **end-to-end data engineering, analytics, and automation workflow** for stock market datasets.
- Designed to showcase **Excel SME, automation, SQL, and Power BI skills** to recruiters and hiring managers.
