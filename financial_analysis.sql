-- Financial Transaction Analysis & Reconciliation
-- Database: MySQL | Project: Finance_Project

CREATE DATABASE IF NOT EXISTS Finance_Project;
USE Finance_Project;

CREATE TABLE IF NOT EXISTS FinancialTransactions (
    TransactionID INT AUTO_INCREMENT PRIMARY KEY,
    TransactionDate DATE,
    TransactionCode VARCHAR(30),
    TransactionAmount DECIMAL(18,2),
    ContractID VARCHAR(30),
    BusinessType VARCHAR(20)
);

CREATE TABLE IF NOT EXISTS ReferenceTransactions (
    ContractID VARCHAR(30) PRIMARY KEY,
    Expected_TransactionCode VARCHAR(30),
    Expected_Amount DECIMAL(18,2),
    Expected_BusinessType VARCHAR(20)
);

-- 1. Basic financial summary
SELECT COUNT(*) AS Total_Transactions,
       SUM(TransactionAmount) AS Total_Transaction_Amount,
       ROUND(AVG(TransactionAmount),2) AS Average_Transaction_Amount,
       MIN(TransactionAmount) AS Minimum_Transaction_Amount,
       MAX(TransactionAmount) AS Maximum_Transaction_Amount
FROM FinancialTransactions;

-- 2. Analysis by transaction code
SELECT TransactionCode, COUNT(*) AS Transaction_Count,
       SUM(TransactionAmount) AS Total_Amount,
       ROUND(AVG(TransactionAmount),2) AS Average_Amount
FROM FinancialTransactions
GROUP BY TransactionCode
ORDER BY Total_Amount DESC;

-- 3. Analysis by business type
SELECT BusinessType, COUNT(*) AS Transaction_Count,
       SUM(TransactionAmount) AS Total_Amount,
       ROUND(AVG(TransactionAmount),2) AS Average_Amount
FROM FinancialTransactions
GROUP BY BusinessType
ORDER BY Total_Amount DESC;

-- 4. Monthly analysis
SELECT DATE_FORMAT(TransactionDate,'%Y-%m') AS Transaction_Month,
       COUNT(*) AS Transaction_Count,
       SUM(TransactionAmount) AS Total_Amount,
       ROUND(AVG(TransactionAmount),2) AS Average_Amount
FROM FinancialTransactions
GROUP BY DATE_FORMAT(TransactionDate,'%Y-%m')
ORDER BY Transaction_Month;

-- 5. High-value transactions
SELECT TransactionID, TransactionDate, TransactionCode,
       TransactionAmount, ContractID, BusinessType
FROM FinancialTransactions
WHERE TransactionAmount > 400000
ORDER BY TransactionAmount DESC;

SELECT COUNT(*) AS High_Value_Transactions,
       SUM(TransactionAmount) AS High_Value_Total
FROM FinancialTransactions
WHERE TransactionAmount > 400000;

-- 6. Top 10 highest-value transactions
SELECT TransactionID, TransactionDate, TransactionCode,
       TransactionAmount, ContractID, BusinessType
FROM FinancialTransactions
ORDER BY TransactionAmount DESC
LIMIT 10;

-- 7. Daily analysis
SELECT TransactionDate, COUNT(*) AS Transaction_Count,
       SUM(TransactionAmount) AS Total_Amount,
       ROUND(AVG(TransactionAmount),2) AS Average_Amount
FROM FinancialTransactions
GROUP BY TransactionDate
ORDER BY TransactionDate;

-- 8. Transaction code distribution
SELECT TransactionCode, COUNT(*) AS Transaction_Count,
       ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (),2) AS Percentage_of_Transactions
FROM FinancialTransactions
GROUP BY TransactionCode
ORDER BY Percentage_of_Transactions DESC;

-- 9. Business type contribution
SELECT BusinessType, SUM(TransactionAmount) AS Total_Amount,
       ROUND(SUM(TransactionAmount) * 100.0 / SUM(SUM(TransactionAmount)) OVER (),2)
       AS Percentage_of_Total_Amount
FROM FinancialTransactions
GROUP BY BusinessType
ORDER BY Total_Amount DESC;

-- 10. Duplicate Contract IDs
SELECT ContractID, COUNT(*) AS Transaction_Count
FROM FinancialTransactions
WHERE ContractID IS NOT NULL
GROUP BY ContractID
HAVING COUNT(*) > 1
ORDER BY Transaction_Count DESC;

-- 11. Null/data-quality check
SELECT SUM(CASE WHEN TransactionDate IS NULL THEN 1 ELSE 0 END) AS Blank_Dates,
       SUM(CASE WHEN TransactionCode IS NULL THEN 1 ELSE 0 END) AS Blank_TransactionCodes,
       SUM(CASE WHEN TransactionAmount IS NULL THEN 1 ELSE 0 END) AS Blank_Amounts,
       SUM(CASE WHEN ContractID IS NULL THEN 1 ELSE 0 END) AS Blank_ContractIDs,
       SUM(CASE WHEN BusinessType IS NULL THEN 1 ELSE 0 END) AS Blank_BusinessTypes
FROM FinancialTransactions;

-- 12. Negative transactions
SELECT TransactionID, TransactionDate, ContractID,
       TransactionCode, TransactionAmount, BusinessType
FROM FinancialTransactions
WHERE TransactionAmount < 0
ORDER BY TransactionAmount;

-- 13. Group Premium Contract ID format validation
-- Required: 5 digits + 5 uppercase letters + P/S/S1
SELECT TransactionID, ContractID, TransactionCode, BusinessType
FROM FinancialTransactions
WHERE TransactionCode = 'Group Premium'
  AND (ContractID IS NULL OR ContractID NOT REGEXP '^[0-9]{5}[A-Z]{5}(P|S|S1)$')
ORDER BY TransactionID;

SELECT COUNT(*) AS Invalid_Group_Contract_IDs
FROM FinancialTransactions
WHERE TransactionCode = 'Group Premium'
  AND (ContractID IS NULL OR ContractID NOT REGEXP '^[0-9]{5}[A-Z]{5}(P|S|S1)$');

-- 14. Populate reference data (run once when reference table is empty)
INSERT INTO ReferenceTransactions
(ContractID, Expected_TransactionCode, Expected_Amount, Expected_BusinessType)
SELECT ContractID, MAX(TransactionCode), MAX(ABS(TransactionAmount)), MAX(BusinessType)
FROM FinancialTransactions
WHERE ContractID IS NOT NULL
GROUP BY ContractID;

-- 15. Create business-type mismatch test cases
SET SQL_SAFE_UPDATES = 0;
UPDATE ReferenceTransactions r
JOIN FinancialTransactions f ON r.ContractID = f.ContractID
SET r.Expected_BusinessType = CASE WHEN f.BusinessType='Group' THEN 'Individual' ELSE 'Group' END
WHERE f.TransactionID > 0 AND MOD(f.TransactionID,100)=0;
SET SQL_SAFE_UPDATES = 1;

SELECT COUNT(*) AS Business_Mismatch_Records
FROM FinancialTransactions f JOIN ReferenceTransactions r ON f.ContractID=r.ContractID
WHERE f.BusinessType <> r.Expected_BusinessType;

-- 16. Create amount mismatch test cases
SET SQL_SAFE_UPDATES = 0;
UPDATE ReferenceTransactions r
JOIN FinancialTransactions f ON r.ContractID = f.ContractID
SET r.Expected_Amount = r.Expected_Amount + 10000
WHERE f.TransactionID > 0 AND MOD(f.TransactionID,200)=0;
SET SQL_SAFE_UPDATES = 1;

SELECT COUNT(*) AS Amount_Mismatch_Records
FROM FinancialTransactions f JOIN ReferenceTransactions r ON f.ContractID=r.ContractID
WHERE f.TransactionAmount <> r.Expected_Amount;

-- 17. Create amount + business-type mismatch cases
SET SQL_SAFE_UPDATES = 0;
UPDATE ReferenceTransactions r
JOIN FinancialTransactions f ON r.ContractID = f.ContractID
SET r.Expected_Amount = r.Expected_Amount + 20000,
    r.Expected_BusinessType = CASE WHEN f.BusinessType='Group' THEN 'Individual' ELSE 'Group' END
WHERE f.TransactionID > 0 AND MOD(f.TransactionID,500)=0;
SET SQL_SAFE_UPDATES = 1;

SELECT COUNT(*) AS Amount_Business_Mismatch
FROM FinancialTransactions f JOIN ReferenceTransactions r ON f.ContractID=r.ContractID
WHERE f.TransactionAmount <> r.Expected_Amount
  AND f.BusinessType <> r.Expected_BusinessType;

-- 18. Create contract-not-found reference cases
INSERT INTO ReferenceTransactions
(ContractID, Expected_TransactionCode, Expected_Amount, Expected_BusinessType)
VALUES
('FAKE_CONTRACT_001','Group Premium',250000,'Group'),
('FAKE_CONTRACT_002','Individual Premium',180000,'Individual'),
('FAKE_CONTRACT_003','Commission',120000,'Individual'),
('FAKE_CONTRACT_004','Group Premium',300000,'Group'),
('FAKE_CONTRACT_005','Individual Premium',220000,'Individual');

-- 19. Reconciliation status classification
SELECT CASE
         WHEN f.ContractID IS NULL THEN 'Contract ID Blank'
         WHEN r.ContractID IS NULL THEN 'Contract Not Found'
         WHEN f.TransactionAmount <> r.Expected_Amount AND f.BusinessType <> r.Expected_BusinessType
              THEN 'Amount & Business Type Mismatch'
         WHEN f.TransactionAmount <> r.Expected_Amount THEN 'Amount Mismatch'
         WHEN f.BusinessType <> r.Expected_BusinessType THEN 'Business Type Mismatch'
         ELSE 'Matched'
       END AS Reconciliation_Status,
       COUNT(*) AS Record_Count
FROM FinancialTransactions f
LEFT JOIN ReferenceTransactions r ON f.ContractID=r.ContractID
GROUP BY Reconciliation_Status
ORDER BY Record_Count DESC;

-- 20. Final exception report
SELECT f.TransactionID, f.TransactionDate, f.ContractID, f.TransactionCode,
       f.TransactionAmount AS Actual_Amount, r.Expected_Amount,
       f.BusinessType AS Actual_BusinessType, r.Expected_BusinessType,
       CASE
         WHEN r.ContractID IS NULL THEN 'Contract Not Found'
         WHEN f.TransactionAmount <> r.Expected_Amount AND f.BusinessType <> r.Expected_BusinessType
              THEN 'Amount & Business Type Mismatch'
         WHEN f.TransactionAmount <> r.Expected_Amount THEN 'Amount Mismatch'
         WHEN f.BusinessType <> r.Expected_BusinessType THEN 'Business Type Mismatch'
         ELSE 'Matched'
       END AS Reconciliation_Status,
       CASE WHEN r.ContractID IS NULL THEN NULL
            ELSE f.TransactionAmount-r.Expected_Amount END AS Amount_Difference
FROM FinancialTransactions f
LEFT JOIN ReferenceTransactions r ON f.ContractID=r.ContractID
WHERE r.ContractID IS NULL
   OR f.TransactionAmount <> r.Expected_Amount
   OR f.BusinessType <> r.Expected_BusinessType
ORDER BY f.TransactionID;

-- 21. Reconciliation KPI summary
SELECT COUNT(*) AS Total_Transactions,
       SUM(CASE WHEN r.ContractID IS NOT NULL
                     AND f.TransactionAmount=r.Expected_Amount
                     AND f.BusinessType=r.Expected_BusinessType THEN 1 ELSE 0 END) AS Matched_Transactions,
       SUM(CASE WHEN r.ContractID IS NULL
                     OR f.TransactionAmount<>r.Expected_Amount
                     OR f.BusinessType<>r.Expected_BusinessType THEN 1 ELSE 0 END) AS Exception_Transactions,
       ROUND(SUM(CASE WHEN r.ContractID IS NOT NULL
                           AND f.TransactionAmount=r.Expected_Amount
                           AND f.BusinessType=r.Expected_BusinessType THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS Match_Rate_Percent,
       ROUND(SUM(CASE WHEN r.ContractID IS NULL
                           OR f.TransactionAmount<>r.Expected_Amount
                           OR f.BusinessType<>r.Expected_BusinessType THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS Exception_Rate_Percent
FROM FinancialTransactions f
LEFT JOIN ReferenceTransactions r ON f.ContractID=r.ContractID;
