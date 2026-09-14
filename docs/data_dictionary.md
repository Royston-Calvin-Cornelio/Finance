\# Data Dictionary



\## Financial Transactions Dataset



The dataset contains financial transaction records used for transaction analysis and reconciliation.



| Column | Data Type | Description | Example |

|---|---|---|---|

| TransactionDate | Date | Date on which the transaction occurred | 2025-04-15 |

| TransactionCode | Text | Classification of the financial transaction | Group Premium |

| TransactionAmount | Decimal | Monetary value associated with the transaction | 250000 |

| ContractID | Text | Identifier used to link transactions with reference data | 54545GLEEAP |

| BusinessType | Text | Business classification of the transaction | Group |



\---



\## TransactionCode



The `TransactionCode` field identifies the type of financial transaction.



| Value | Meaning |

|---|---|

| Group Premium | Premium transaction associated with group business |

| Individual Premium | Premium transaction associated with individual business |

| Commission | Commission-related transaction |



\---



\## BusinessType



The `BusinessType` field identifies the business classification.



| Value | Meaning |

|---|---|

| Group | Transaction associated with group business |

| Individual | Transaction associated with individual business |



\---



\## ContractID



`ContractID` is used as the primary business identifier for reconciliation.



For Group Premium transactions, the expected format is:



```text

5 digits + 5 letters + life indicator

EG: 54545GLEEAP

54545 | GLEEA | P
  |       |     |
  |       |     └── Life Indicator
  |       └──────── 5-character identifier
  └──────────────── 5-digit identifier

