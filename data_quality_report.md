# Customer Churn — Data Quality and Exploratory Analysis

## 1. Objective
Audit and clean the supplied customer churn dataset, document quality decisions, and summarize initial patterns relevant to customer churn.

## 2. Dataset overview
- Records: 15
- Columns: 11
- Duplicate full rows: 0
- Duplicate CustomerID values: 0
- Missing values: 0
- Churn = Yes: 7 (46.7%)
- Churn = No: 8 (53.3%)

## 3. Data-quality checks and decisions
1. **Column structure and data types:** Checked all 11 columns. CustomerID and category fields are text; Age, TenureMonths, SupportTickets are integer-valued; MonthlyCharges and TotalCharges are numeric decimals.
2. **Missing values:** No missing values were found in the supplied file. No imputation was needed.
3. **Duplicates:** No duplicate full rows or duplicate CustomerID values were found. No rows were removed.
4. **Text consistency:** Trimmed leading/trailing whitespace from text columns. No category values required substantive recoding.
5. **Numeric validation:** Values are non-negative for Age, TenureMonths, MonthlyCharges, TotalCharges and SupportTickets. No impossible values were observed in this sample.
6. **Outliers:** Applied the 1.5 × IQR rule to each numeric column. The rule flagged no outliers. This does not prove all values are valid; domain checks remain important.
7. **Charge consistency:** Compared TotalCharges with MonthlyCharges × TenureMonths. Values agree to normal floating-point rounding (within 0.01) for all rows. This is a consistency check, not proof of billing correctness.

## 4. Summary statistics
        Column  count        mean         std   min    25%     50%     75%     max
           Age   15.0   39.066667   11.392144 23.00  30.00   38.00   47.00   60.00
  TenureMonths   15.0   18.800000   14.333776  2.00   7.00   15.00   27.00   48.00
MonthlyCharges   15.0   84.656667   42.739521 49.99  49.99   79.99  114.99  149.99
  TotalCharges   15.0 2088.478667 2407.365767 99.98 449.91 1099.78 3209.73 7199.52
SupportTickets   15.0    2.533333    1.922300  0.00   1.00    2.00    4.00    6.00

## 5. Initial exploratory findings
- The observed churn rate is **46.7%** (7 of 15 customers).
- Average tenure is **18.8 months** overall.
- Customers marked as churned have mean tenure **7.0 months**, compared with **29.1 months** for non-churned customers.
- Mean support tickets are **4.29** for churned customers versus **1.00** for non-churned customers.
- In this sample, all Month-to-Month customers are marked churned; all One Year and Two Year customers are marked not churned.
- Basic-plan customers have more churned records than non-churned records in this sample (5 versus 2).
- These are **descriptive patterns only**. The dataset has just 15 records, so results are small-sample observations and should not be treated as causal or generalized to all customers.

## 6. Cleaning log
| Check | Result | Decision |
|---|---|---|
| Missing values | None | No imputation |
| Duplicate rows | 0 | No removal |
| Duplicate customer IDs | 0 | No removal |
| IQR outliers | None flagged | No outlier removal |
| Charge formula check | Consistent within 0.01 | No correction |
| Text whitespace | Trimmed in text fields | Kept all records |

## 7. Limitations and next steps
- Only 15 records are available; patterns may be unstable.
- The contract/churn relationship is perfectly separated in this small sample, which may reflect how the synthetic data was generated.
- Next steps for a larger real dataset: validate business rules with a data owner, investigate churn by tenure and contract, visualize distributions, and evaluate models only with a sufficiently large dataset and proper train/test separation.

## 8. Deliverables
- `customer_churn_cleaned.csv`: cleaned dataset.
- `customer_churn_analysis.xlsx`: workbook with cleaned data, data profile, numeric summary, IQR checks, churn breakdowns, and cleaning log.
- This report: data-quality decisions and initial patterns.
