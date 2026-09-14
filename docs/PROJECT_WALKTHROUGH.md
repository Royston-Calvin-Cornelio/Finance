\# Project Walkthrough



\## 1. Problem Statement



Financial organizations process large volumes of transactions that need to be validated and reconciled against expected or reference data.



Manual reconciliation can be time-consuming and may make it difficult to identify exceptions consistently.



The objective of this project was to build an end-to-end workflow that could:



\- Analyze financial transactions

\- Validate data quality

\- Compare actual data with reference data

\- Identify discrepancies

\- Classify reconciliation exceptions

\- Calculate reconciliation KPIs

\- Automate exception reporting



\---



\# 2. Dataset



The project uses a dataset containing:



\*\*1,000,000 financial transactions\*\*



\### Key fields



| Field | Description |

|---|---|

| TransactionDate | Transaction date |

| TransactionCode | Type of transaction |

| TransactionAmount | Transaction amount |

| ContractID | Contract/reference identifier |

| BusinessType | Group or Individual |



\### Transaction Types



\- Group Premium

\- Individual Premium

\- Commission



\---



\# 3. Data Analysis Using Python



Python and Pandas were used to analyze the transaction dataset.



\### Analysis performed



\- Record count

\- Total transaction amount

\- Average transaction amount

\- Minimum and maximum transaction amount

\- Transaction analysis by transaction type

\- Business-type analysis

\- Monthly transaction analysis

\- High-value transaction identification

\- Duplicate detection

\- Missing-value analysis

\- Life Indicator analysis

\- Contract ID validation



\### Results



| Metric | Result |

|---|---:|

| Total Records | 1,000,000 |

| Total Transaction Amount | 252,514,880,709 |

| Average Transaction Amount | 252,514.88 |

| Minimum Amount | 5,000 |

| Maximum Amount | 500,000 |

| Duplicate Records | 0 |

| Missing Values | 0 |

| Transactions > 400,000 | 202,533 |



\---



\# 4. SQL Analysis



MySQL was used to perform financial analysis and validation.



The SQL layer included:



\- Transaction analysis

\- Business-type analysis

\- Monthly analysis

\- High-value transaction analysis

\- Duplicate Contract ID detection

\- Missing-value checks

\- Negative transaction checks

\- Invalid Contract ID checks

\- Contract ID format validation



SQL was also used to build the reconciliation logic.



\---



\# 5. Reference Data



A reference dataset was created containing expected information for financial transactions.



The reference data included:



\- Contract ID

\- Expected transaction code

\- Expected amount

\- Expected business type



The actual transaction data was compared against the reference data using:



\*\*ContractID\*\*



\---



\# 6. Reconciliation Process



The reconciliation process compares actual transaction information with expected reference information.



\### Comparison logic



```text

Actual Transaction

&#x20;      |

&#x20;      v

Find ContractID in Reference Data

&#x20;      |

&#x20;      v

Compare Amount

&#x20;      |

&#x20;      v

Compare Business Type

&#x20;      |

&#x20;      v

Assign Reconciliation Status

