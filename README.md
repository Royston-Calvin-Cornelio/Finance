# Financial Transaction Analysis & Automated Reconciliation

## 📌 Project Overview

An end-to-end finance analytics and reconciliation project designed to simulate a real-world financial operations environment.

The project analyzes **1 million financial transactions**, performs data-quality validation, reconciles transactions against reference data, identifies exceptions, and automates reconciliation reporting using **Python, SQL, Excel and VBA**.

---

## 🎯 Business Objective

The objective of this project is to demonstrate how finance and accounting knowledge can be combined with data analytics and automation to solve practical business problems.

The project focuses on:

- Financial transaction analysis
- Data reconciliation
- Data-quality validation
- Exception identification
- Financial reporting
- Process automation
- Management reporting

---

## 🛠️ Tools & Technologies

| Technology | Purpose |
|---|---|
| Python | Large-scale transaction analysis |
| Pandas | Data manipulation and analysis |
| MySQL | SQL-based financial analysis and reconciliation |
| Microsoft Excel | Reconciliation and reporting |
| VBA | Reconciliation automation |
| Power BI | Dashboarding and visualization |

---

# 📂 Project Structure

```text
Finance
│
├── python
│   ├── financial_analysis
│   └── financial_transactions.csv
│
├── sql
│   └── financial_analysis.sql
│
├── Excel
│   ├── Financial_Reconciliation_Automation.bas
│   └── Reconciliation_Examples.csv
│
├── screenshots
│   ├── reconciliation_summary.png
│   └── reconciliation_exceptions.png
│
├── README.md
└── .gitignore
```

> The original Excel `.xlsm` workbook is maintained locally because of GitHub file-size limitations. The VBA source code and reconciliation examples are included in the repository.

---

# 🐍 Python Analysis

Python and Pandas were used to analyze **1,000,000 financial transactions**.

### Analysis Performed

- Transaction volume analysis
- Total transaction amount
- Average transaction amount
- Minimum and maximum transaction values
- Transaction analysis by transaction code
- Business-type analysis
- Monthly transaction trends
- High-value transaction identification
- Duplicate detection
- Missing-value analysis
- Life-indicator analysis
- Contract ID format validation

### Python Results

| Metric | Result |
|---|---:|
| Total Records | 1,000,000 |
| Total Transaction Amount | 252,514,880,709 |
| Average Transaction Amount | 252,514.88 |
| Maximum Transaction Amount | 500,000 |
| Minimum Transaction Amount | 5,000 |
| Duplicate Records | 0 |
| Missing Values | 0 |

### Transaction Analysis

The dataset contains three major transaction categories:

- Group Premium
- Individual Premium
- Commission

The analysis calculates transaction count, total transaction amount and average transaction amount for each category.

### Monthly Analysis

Python was also used to analyze transaction volumes and transaction amounts by month, helping identify trends across the financial year.

---

# 🗄️ SQL Analysis

MySQL was used to perform financial analysis, data-quality checks and reconciliation analysis.

### SQL Analysis Includes

- Financial summary
- Transaction analysis by transaction code
- Business-type analysis
- Monthly transaction analysis
- High-value transaction analysis
- Top transaction analysis
- Duplicate Contract ID detection
- NULL-value checks
- Negative transaction detection
- Contract ID format validation
- Reconciliation status classification
- Exception reporting
- Reconciliation KPI calculation

The complete SQL analysis is available in:

```text
sql/financial_analysis.sql
```

---

# 🔄 Financial Reconciliation

The reconciliation process compares actual financial transaction data against reference data.

### Reconciliation Process

```text
Transaction Data
       ↓
Reference Data
       ↓
Contract ID Matching
       ↓
Amount Comparison
       ↓
Business Type Comparison
       ↓
Reconciliation Status
       ↓
Exception Report
```

### Reconciliation Logic

Each transaction is compared with the corresponding reference record using **Contract ID**.

The process validates:

1. Contract ID availability
2. Contract ID existence in reference data
3. Transaction amount
4. Business type
5. Overall reconciliation status

---

# ⚠️ Reconciliation Categories

The reconciliation process classifies transactions into the following categories:

### ✅ Matched

Actual transaction data agrees with the reference data.

### ⚠️ Amount Mismatch

The transaction amount differs from the expected reference amount.

### ⚠️ Business Type Mismatch

The business type differs from the expected reference value.

### ⚠️ Amount & Business Type Mismatch

Both the transaction amount and business type differ from the reference data.

### ❌ Contract Not Found

The transaction Contract ID does not exist in the reference data.

---

# 📊 Reconciliation Results

The Excel reconciliation tool processed **1,000,000 transactions**.

| Metric | Result |
|---|---:|
| Total Transactions | 1,000,000 |
| Matched Transactions | 727,850 |
| Exception Transactions | 272,150 |
| Match Rate | 73% |
| Exception Rate | 27% |
| Duplicate Contract IDs | 162,741 |

### Exception Breakdown

| Exception Type | Records |
|---|---:|
| Amount Mismatch | 267,483 |
| Business Type Mismatch | 1,000 |
| Amount & Business Type Mismatch | 1,737 |
| Contract Not Found | 1,930 |

---

# 📗 Excel & VBA Automation

An Excel-based reconciliation tool was developed using VBA to automate repetitive reconciliation activities.

### VBA Automation Performs

- Reconciliation calculation
- Data-quality checks
- Exception identification
- Reconciliation status classification
- Amount difference calculation
- Exception-rate calculation
- Automated exception report generation

The VBA source code is available here:

```text
Excel/Financial_Reconciliation_Automation.bas
```

The automation transforms a manual reconciliation process into a repeatable workflow.

---

# 📸 Reconciliation Screenshots

## Reconciliation Summary

![Reconciliation Summary](screenshots/reconciliation_summary.png)

## Reconciliation Exceptions

![Reconciliation Exceptions](screenshots/reconciliation_exceptions.png)

---

# 🔍 Data Quality Controls

The project includes multiple data-quality controls designed to identify common financial-data issues.

### Controls Include

- Duplicate transaction detection
- Duplicate Contract ID detection
- Missing Contract ID detection
- Missing transaction amount detection
- Negative transaction detection
- Invalid Contract ID detection
- Contract ID format validation
- Reference-data validation
- Amount mismatch detection
- Business type mismatch detection

These controls help identify potential issues before financial reporting or downstream processing.

---

# 💼 Finance Skills Demonstrated

This project demonstrates practical knowledge of:

- Financial reconciliation
- Financial transaction analysis
- Financial reporting
- Data-quality controls
- Exception management
- Reference-data validation
- Transaction-level investigation
- Financial data analysis
- Automated reporting

---

# 💻 Technical Skills Demonstrated

### Programming & Data

- Python
- Pandas
- SQL
- MySQL

### Excel & Automation

- Microsoft Excel
- Advanced Excel formulas
- XLOOKUP
- VBA
- Process automation
- Exception reporting

### Business Intelligence

- Power BI

---

# 🚀 Business Value

This project demonstrates the ability to combine **finance/accounting knowledge with technical and analytical skills** to automate repetitive processes and improve financial data analysis.

The project is particularly relevant to roles involving:

- Financial Data Analyst
- Finance Analyst
- Reconciliation Analyst
- MIS Analyst
- Reporting Analyst
- Business Analyst
- Data Analyst
- Finance Operations Analyst

---

# 📈 Key Project Highlights

- Analyzed **1 million financial transactions**
- Built financial analysis using **Python and Pandas**
- Developed multiple **MySQL financial analysis queries**
- Designed a transaction reconciliation process
- Automated reconciliation using **Excel VBA**
- Created automated exception reporting
- Implemented data-quality validation controls
- Identified and categorized reconciliation exceptions
- Combined finance knowledge with data analytics and automation

---

# 👨‍💼 About

M.Com graduate with finance and financial reporting experience, combining accounting knowledge with technical skills in:

**Python | SQL | Excel | VBA | Power BI | Data Analysis**

This project demonstrates the practical application of these skills in a financial-data and reconciliation environment.

---

## ⭐ Project Objective

The overall objective of this project is to demonstrate the ability to take a large financial dataset, analyze it, validate its quality, reconcile it against reference information, identify exceptions and automate the reporting process.

```text
Large Financial Dataset
        ↓
Python Analysis
        ↓
SQL Analysis
        ↓
Data Quality Checks
        ↓
Financial Reconciliation
        ↓
VBA Automation
        ↓
Exception Reporting
```