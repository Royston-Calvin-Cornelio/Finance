# Financial Transaction Analysis & Automated Reconciliation

### Python | SQL | Excel | VBA | Pandas | MySQL

An end-to-end financial data analysis and reconciliation project designed to simulate real-world finance operations involving transaction analysis, data-quality validation, financial reconciliation, exception identification, and automated reporting.

---

## 📌 Project Overview

Financial organizations process large volumes of transactions that need to be validated and reconciled against reference data.

This project demonstrates how a finance professional can combine **financial knowledge with technical skills** to analyze transaction data, identify discrepancies, perform data-quality checks, and automate reconciliation reporting.

The project processes **1 million financial transactions** and uses multiple technologies across the workflow:

- Python & Pandas for data analysis
- MySQL for financial analysis and reconciliation queries
- Excel for reconciliation and reporting
- VBA for automation
- GitHub for version control and project documentation

---

## 🎯 Business Objective

The main objective is to build a practical financial reconciliation workflow that can:

1. Analyze large transaction datasets
2. Identify transaction patterns
3. Detect data-quality issues
4. Compare actual transactions against reference data
5. Identify financial discrepancies
6. Classify reconciliation exceptions
7. Calculate reconciliation KPIs
8. Automate exception reporting using VBA

This simulates activities commonly performed in **financial reporting, accounting operations, insurance finance, reconciliation, and data-analysis functions**.

---

# 📊 Dataset

The primary dataset contains **1,000,000 financial transaction records**.

### Main fields

| Column | Description |
|---|---|
| TransactionDate | Date of the transaction |
| TransactionCode | Type of financial transaction |
| TransactionAmount | Monetary value of the transaction |
| ContractID | Unique/reference contract identifier |
| BusinessType | Group or Individual business |

### Transaction Types

- Group Premium
- Individual Premium
- Commission

### Business Types

- Group
- Individual

---

# 🐍 Python Analysis

Python and Pandas were used to perform exploratory financial analysis and data-quality checks on the 1 million transaction dataset.

### Analysis performed

- Total transaction count
- Total transaction amount
- Average transaction amount
- Minimum and maximum transaction values
- Transaction analysis by transaction type
- Business-type analysis
- Monthly transaction analysis
- High-value transaction identification
- Duplicate detection
- Missing-value analysis
- Life Indicator analysis
- Contract ID format validation

### Key Python Results

| Metric | Result |
|---|---:|
| Total Records | 1,000,000 |
| Total Transaction Amount | 252,514,880,709 |
| Average Transaction Amount | 252,514.88 |
| Minimum Transaction Amount | 5,000 |
| Maximum Transaction Amount | 500,000 |
| Duplicate Records | 0 |
| Missing Values | 0 |
| Transactions > 400,000 | 202,533 |

### Transaction Analysis

| Transaction Type | Records | Total Amount |
|---|---:|---:|
| Group Premium | 449,489 | 113,549,309,361 |
| Individual Premium | 400,399 | 101,127,936,162 |
| Commission | 150,112 | 37,837,635,186 |

---

# 🗄️ SQL Analysis

MySQL was used to perform financial analysis, data-quality validation, duplicate detection, and reconciliation analysis.

### SQL analysis includes

1. Total transaction analysis
2. Transaction-type analysis
3. Business-type analysis
4. Monthly transaction analysis
5. High-value transaction analysis
6. Duplicate Contract ID detection
7. Missing-value checks
8. Negative transaction detection
9. Invalid Contract ID detection
10. Contract ID format validation
11. Reference-data reconciliation
12. Exception classification
13. Amount difference calculation
14. Reconciliation KPI calculation

---

# 🔄 Financial Reconciliation

A separate reference dataset was created to simulate expected financial transaction information.

The actual transaction data was compared against the reference data using:

**ContractID**

The reconciliation checks:

- Transaction amount
- Business type
- Contract existence

---

## Reconciliation Logic

Each transaction is classified into one of the following categories:

### ✅ Matched

Actual amount and business type match the reference data.

### ⚠️ Amount Mismatch

Actual transaction amount differs from the expected amount.

### ⚠️ Business Type Mismatch

Actual business type differs from the expected business type.

### 🔴 Amount & Business Type Mismatch

Both the transaction amount and business type differ.

### ❌ Contract Not Found

The Contract ID does not exist in the reference dataset.

---

# 📈 Reconciliation Results

The Excel reconciliation tool produced the following results:

| Reconciliation Status | Records |
|---|---:|
| Matched | 727,850 |
| Amount Mismatch | 267,483 |
| Business Type Mismatch | 1,000 |
| Amount & Business Type Mismatch | 1,737 |
| Contract Not Found | 1,930 |
| **Total Transactions** | **1,000,000** |

### Key KPIs

**Match Rate:** 73%

**Exception Rate:** 27%

**Total Exceptions:** 272,150

**Total Amount Difference:** -70,231,929

---

# 📊 Excel Reconciliation Tool

Excel was used to create a reconciliation workbook containing:

### Transaction_Data

Contains the transaction dataset.

### Reference_Data

Contains expected transaction information.

### Reconciliation

Performs record-level comparison between actual and expected data.

### Summary

Provides reconciliation KPIs and data-quality indicators.

---

# ⚙️ VBA Automation

VBA was used to automate the reconciliation workflow.

The automation performs the following steps:

1. Recalculates reconciliation formulas
2. Updates summary metrics
3. Performs data-quality checks
4. Calculates exception rate
5. Records the last reconciliation run
6. Exports reconciliation exceptions
7. Generates a timestamped exception report

This reduces manual reconciliation work and creates a repeatable reporting process.

---

# 🔍 Data Quality Controls

The project includes multiple financial data-quality controls.

### Controls implemented

- Duplicate record detection
- Duplicate Contract ID detection
- Blank Contract ID detection
- Blank transaction amount detection
- Negative transaction detection
- Invalid Contract ID detection
- Contract ID format validation
- Missing reference records
- Transaction amount mismatch detection
- Business type mismatch detection

These controls demonstrate how automated validation can be incorporated into financial data workflows.

---

# 📁 Project Structure

```text
Finance/
│
├── python/
│   ├── financial_analysis
│   └── financial_transactions.csv
│
├── sql/
│   └── financial_analysis.sql
│
├── Excel/
│   ├── Financial_Reconciliation_Automation.bas
│   └── Reconciliation_Examples.csv
│
├── screenshots/
│   ├── reconciliation_summary.png
│   └── reconciliation_exceptions.png
│
├── README.md
├── requirements.txt
└── .gitignore