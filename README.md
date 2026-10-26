# Customer Churn: Data Quality and Exploratory Analysis

## Objective

This project profiles and cleans a customer churn dataset, documents data-quality decisions, and presents descriptive exploratory analysis of customer churn patterns.

## Dataset

The cleaned dataset contains 15 customer records and 11 fields:

- Identifiers and demographics: `CustomerID`, `Gender`, `Age`
- Customer relationship: `TenureMonths`, `SubscriptionType`, `ContractType`, `SupportTickets`, `PaymentMethod`
- Charges: `MonthlyCharges`, `TotalCharges`
- Outcome: `Churn`

## Methodology

The project deliverables document these steps:

1. Profile columns, data types, missing values, duplicate rows/IDs, and numeric outliers using the 1.5 × IQR rule.
2. Apply conservative cleaning: trim text whitespace; retain all rows because no missing values or duplicates were identified; do not remove values solely because of the IQR check.
3. Validate non-negative numeric values and compare `TotalCharges` against `MonthlyCharges × TenureMonths` (within 0.01).
4. Summarize numeric variables and compare churn across contract, subscription, tenure, and support-ticket groups.

## Findings

- 15 records and 11 columns were analyzed.
- The supplied analysis reports no missing values, duplicate full rows, duplicate customer IDs, or IQR-flagged numeric outliers.
- 7 of 15 customers (46.7%) are marked as churned.
- In this sample, churned customers have lower mean tenure (7.0 months) and more mean support tickets (4.29) than non-churned customers (29.1 months and 1.00 ticket).
- All month-to-month customers in the sample are marked as churned; all one- and two-year customers are marked as not churned.

## Limitations

This is a 15-record dataset. All findings are descriptive only and must not be interpreted as causal or generalized to a larger customer population. The contract/churn pattern is perfectly separated in this sample and may reflect synthetic-data construction. The charge calculation check supports internal consistency only; it does not prove billing accuracy.

## Repository contents

| Path | Description |
| --- | --- |
| `data/customer_churn_cleaned.csv` | Cleaned CSV dataset used for the analysis. |
| `analysis/customer_churn_analysis.xlsx` | Workbook with cleaned data, data profile, numeric summary, IQR checks, churn/category summaries, group means, and cleaning log. |
| `reports/data_quality_report.md` | Detailed data-quality decisions, validation checks, summary statistics, findings, and limitations. |

## Reproducibility

Open the workbook for the documented tables and checks, or use the CSV for independent analysis. No code or raw source dataset was included in the submitted ZIP, so this repository does not claim to reproduce the cleaning process from raw input.
