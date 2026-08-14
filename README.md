# Scrap Sale Analysis Dashboard

An Excel dashboard analyzing scrap sale transaction data — built using pivot tables, pivot charts, and Excel formulas (SUMIFS, COUNTIFS, AVERAGE, SUBTOTAL) to summarize revenue, transactions, and buyer performance over time.

## 📊 Data Range

- **Period covered:** 26-Jan-2022 to 23-Jul-2025
- **Source:** Scrap_Sale transaction log

## 📁 Workbook Structure

| Sheet | Description |
|---|---|
| **DashBoard** | Main visual dashboard — KPI summary, tables, and charts |
| **Sheet2** | Pivot tables powering part of the dashboard (Top Buyers, Monthly Revenue Trend) |
| **Scrap_Sale** | Raw transaction data (invoice-level records) |
| **Item_Codes** | Reference table mapping item codes to categories |

## 📈 Dashboard Overview

### Key Metrics (KPI cards)
- Total Revenue (Rs.)
- Total Transactions
- Unique Buyers
- Average Transaction Value (Rs.)
- VAT Collected

### Tables & Breakdowns
- **Revenue by Category** — Cardboard, Polymers, Metal, Oil, Wooden, Electronic, E-Waste, N/A
- **Revenue by Hazard Status** — Hazard vs Non-Hazard split
- **Revenue by Disposed Method** — Recycle, Reuse, N/A
- **Top 10 Buyers by Revenue** — ranked by total revenue and transaction count (pivot table)
- **Monthly Revenue Trend** — revenue grouped by Year and Month, Jan-22 through Jul-25 (pivot table)

### Charts
- Revenue by Category (bar chart)
- Revenue by Disposal Method (pie chart)
- Top 10 Buyers by Revenue (bar chart)
- Monthly Revenue Trend (line chart)
- Revenue Share: Hazard Status (pie chart)

## 🧮 Formulas Used

The KPI cards and category/hazard/disposal-method tables on the DashBoard sheet are built with live Excel formulas (not static pivot values), so they recalculate automatically as the underlying `Scrap_Sale` data changes.

**KPI Cards**
```excel
Total Revenue          = SUM(Scrap_Sale!N2:N1099)
Total Transactions     = SUBTOTAL(103, Scrap_Sale[INVOICE NO])
Avg Transaction Value  = AVERAGE(Scrap_Sale[TOTAL AMOUNT RS.])
VAT Collected          = SUM(Scrap_Sale[VAT])
```

**Revenue by Category table**
```excel
Revenue (Rs.)   = SUMIFS(Scrap_Sale!$N$2:$N$1099, Scrap_Sale!$I$2:$I$1099, $B12)
Transactions    = COUNTIFS(Scrap_Sale!$I$2:$I$1099, $B12)
Avg Value (Rs.) = IFERROR(C12/D12, 0)
% of Total      = C12/$B$6
```

**Revenue by Hazard Status table**
```excel
Revenue (Rs.) = SUMIFS(Scrap_Sale!$N$2:$N$1099, Scrap_Sale!$Q$2:$Q$1099, $H12)
% of Total    = I12/$B$6
```

**Revenue by Disposed Method table**
```excel
Revenue (Rs.) = SUMIFS(Scrap_Sale[TOTAL AMOUNT RS.], Scrap_Sale[Disposed Method], H17)
% of Total    = I17/$B$6
```

**Top 10 Buyers & Monthly Revenue Trend (Sheet2)**

Built with native Excel **PivotTable Value Field Settings** rather than formulas:

| Metric | Calculation Type |
|---|---|
| Sum of Revenue | Sum |
| Monthly grouping | Group by Year + Month |

## 🛠️ Tools Used

- Microsoft Excel
- Excel Tables & structured references (`Scrap_Sale[...]`)
- SUMIFS / COUNTIFS / AVERAGE / SUBTOTAL / IFERROR
- PivotTables & PivotCharts

## ⚠️ Note

This dataset was used for training/practice purposes. Buyer names and contact details included are part of the original sample dataset.
