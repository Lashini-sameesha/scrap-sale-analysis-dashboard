# Scrap Sale Analysis Dashboard

An Excel dashboard that analyses scrap sale transactions. It uses cleaned data, formulas, pivot tables and charts to show revenue, transactions and buyer performance over time.

## Data
- **Period:** 26-Jan-2022 to 23-Jul-2025
- **Records:** 1,096 invoice lines (after cleaning)
- **File:** `Scrap_Sale_Analysis_Dashboard_Updated.xlsx`

## Data Cleaning
- Renamed messy column headers to simple names (for example, "TOTAL AMOUNT RS." became "Total Amount")
- Removed extra rows (1,098 reduced to 1,096)
- Converted the data into an Excel Table (`Scrap_Sale2`) so formulas use clear structured references

## Workbook Structure
| Sheet | Description |
|---|---|
| Dashboard | KPI cards, summary tables and charts |
| Pivots | Pivot tables for Top 10 Buyers and Monthly Revenue Trend |
| Scrap_Sale (2) | Cleaned transaction data (used by the dashboard) |
| Scrap_Sale | Original raw data |
| Item_Codes | Reference table that maps item codes to categories |

## Dashboard Contents
**KPI cards:** Total Revenue, Total Transactions, Unique Buyers, Average Transaction Value, VAT Collected

**Tables:** Revenue by Category, Revenue by Hazard Status, Revenue by Disposed Method, Top 10 Buyers

**Charts:**
- Revenue by Category (bar chart)
- Revenue by Disposed Method (pie chart)
- Revenue by Hazard Status (pie chart)
- Top 10 Buyers by Revenue (bar chart)
- Monthly Revenue Trend (line chart)

## Formulas and Tools Used
- `SUMIFS`, `COUNTIFS`, `AVERAGE`, `SUBTOTAL`, `IFERROR`, `ROUND`
- Excel Tables and structured references
- PivotTables and PivotCharts (grouped by year and month)

Example:
```
Revenue by Category = SUMIFS(Scrap_Sale2[Total Amount], Scrap_Sale2[Category], $B12)
```

## Key Insights
- Total revenue is about **Rs. 22.3 million** from 1,096 transactions and 207 unique buyers.
- **Cardboard (69%)** and **Polymers (28%)** together make up 97% of revenue.
- The top buyer, **K Maheshwaran**, contributes about **51%** of total revenue, so the business depends heavily on one buyer.
- **Non-Hazard** items bring 98.9% of revenue.
- **Recycle** brings 83% of revenue, while **Reuse** has more transactions (728 vs 368).
- **2024** was the best year (about Rs. 8.3 million), and the highest month was **March 2024**.

## Note
This dataset was used for training/practice purposes. Buyer names and contact details included are part of the original sample dataset.
