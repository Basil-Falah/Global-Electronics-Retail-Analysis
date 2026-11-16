# Global Electronics Retail — Data Analytics Project

## Project Overview
An end-to-end retail data analytics project built in Microsoft Excel. The project combines data preparation, lookup-based calculated fields, PivotTables, a Data Model, and an interactive dashboard to analyze sales, revenue, gross profit, orders, products, customers, stores, and delivery performance.

## Project Files
- `Analysis/Global_Electronics_Analysis.xlsx` — the working analysis workbook. It contains the source tables, calculated columns, Power Query connections, the Excel Data Model, and PivotTables.
- `Dashboard/Global_Electronics_Dashboard.xlsx` — the presentation workbook containing the interactive dashboard and its supporting calculation sheet.
- `docs/FORMULAS_AND_DATA_MODEL.md` — formula logic, table structure, and validation notes.

## Workbook Structure
The analysis workbook contains five worksheets:
- **Sales** — 62,884 transaction-line records.
- **Products** — 2,517 product records.
- **Customers** — 15,266 customer records.
- **Stores** — 67 store records.
- **Pivot Table** — five PivotTables.

The workbook also contains four Power Query connections (`Sales`, `Products`, `Customers`, and `Stores`) and an Excel Data Model.

## Tools and Techniques
- Microsoft Excel
- Power Query connections
- Excel Tables and structured references
- `XLOOKUP`
- Calculated columns and KPI logic
- PivotTables and the Excel Data Model
- Interactive dashboard filters and charts

## Core Metrics
- Revenue
- Total Cost
- Gross Profit
- Profit Margin
- Order Count
- Units Sold
- Delivery Performance

See `docs/FORMULAS_AND_DATA_MODEL.md` for the exact calculated-column formulas and their business meaning.

## Important Metric Definition
The calculated unit price and unit cost are retrieved from the product table in **USD**. Therefore, the calculated Revenue, Cost, and Gross Profit are USD-based values. `Currency Code` is retained as a transaction attribute; it is not used by the listed formulas to convert the financial amounts into another currency.

## Validation Notes
The supplied analysis workbook was inspected without saving or modifying it. The following checks were performed against its stored formula results:
- 62,884 Sales rows were present.
- All 10 calculated columns in Sales contained formulas across the transaction rows.
- No `Not Found` lookup results were detected in those calculated columns.
- Revenue, Gross Profit, and Profit Margin cached results matched their formula logic within the tested precision.
- The `Store State` values were compared with the Stores table using `StoreKey`; no mismatches were found.

## How to Use
1. Open the analysis workbook in a recent version of Microsoft Excel.
2. Use the Analysis workbook to inspect source tables, calculated columns, and PivotTables.
3. Open the Dashboard workbook to explore the interactive visual analysis.

**Compatibility:** The analysis workbook uses `XLOOKUP`, which requires a recent Excel version (Microsoft 365 or Excel 2021 and later).
