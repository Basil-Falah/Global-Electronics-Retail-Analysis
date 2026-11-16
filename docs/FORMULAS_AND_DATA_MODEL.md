# Formulas, Data Model, and Validation

## 1. Source Tables

| Table | Rows (excluding header) | Role |
|---|---:|---|
| `Sales` | 62,884 | Transaction-line fact table |
| `Products` | 2,517 | Product lookup/dimension |
| `Customers` | 15,266 | Customer lookup/dimension |
| `Stores` | 67 | Store lookup/dimension |

The workbook contains four Power Query connections named `Sales`, `Products`, `Customers`, and `Stores`, plus an Excel Data Model and five PivotTables on the `Pivot Table` worksheet.

## 2. Sales Calculated Columns

The following are the formulas found in the uploaded workbook. Structured references are retained as stored.

| Sales column | Formula | Purpose |
|---|---|---|
| `Product Name` | `=XLOOKUP(Sales[[#This Row],[ProductKey]],Products[ProductKey],Products[Product Name],"Not Found")` | Retrieves the product name using `ProductKey`. |
| `Gender` | `=XLOOKUP(Sales[[#This Row],[CustomerKey]],'Customers'!A:A,'Customers'!B:B,"Not Found")` | Retrieves customer gender using `CustomerKey`. |
| `Customer Name` | `=XLOOKUP(Sales[[#This Row],[CustomerKey]],'Customers'!A:A,'Customers'!C:C,"Not Found")` | Retrieves customer name using `CustomerKey`. |
| `Country` | `=XLOOKUP(Sales[[#This Row],[CustomerKey]],'Customers'!A:A,'Customers'!H:H,"Not Found")` | Retrieves the customer's country. This is the customer-market field. |
| `Unit Price` | `=XLOOKUP(Sales[[#This Row],[ProductKey]],Products[ProductKey],Products[Unit Price USD],"Not Found")` | Retrieves unit selling price in USD. |
| `Total Revenue` | `=Sales[[#This Row],[Quantity]]*Sales[[#This Row],[Unit Price]]` | Quantity multiplied by unit selling price. |
| `Total Cost` | `=Sales[[#This Row],[Quantity]]*XLOOKUP(Sales[[#This Row],[ProductKey]],Products[ProductKey],Products[Unit Cost USD],"Not Found")` | Quantity multiplied by product unit cost in USD. |
| `Gross Profit` | `=Sales[[#This Row],[Total Revenue]]-Sales[[#This Row],[Total Cost]]` | Revenue minus cost. |
| `Profit Margin` | `=Sales[[#This Row],[Gross Profit]]/Sales[[#This Row],[Total Revenue]]` | Gross profit divided by revenue at the transaction-line level. |
| `Store Country` | `=XLOOKUP(Sales[[#This Row],[StoreKey]],Stores[StoreKey],Stores[Country],"Not Found")` | Retrieves the store's country using `StoreKey`. |

Each of the 10 calculated columns above contains a formula for all 62,884 Sales rows.

## 3. Additional Store Attribute

`Store State` is present in the Sales table as stored values rather than an Excel formula in the inspected file. Its values were checked against the `Stores` table using `StoreKey`; no mismatches were found in the 62,884 rows. It was left unchanged.

## 4. Metric Interpretation

- **Revenue:** `Quantity × Unit Price`, where Unit Price is looked up from `Products[Unit Price USD]`.
- **Total Cost:** `Quantity × Unit Cost USD`.
- **Gross Profit:** `Revenue − Total Cost`.
- **Profit Margin:** `Gross Profit ÷ Revenue`.

The workbook's calculated financial amounts are USD-based. The `Currency Code` field is present in Sales but is not referenced by these formulas for currency conversion. Do not describe these calculated amounts as converted transaction-currency totals.

When reporting an overall profit margin, calculate it as **total gross profit ÷ total revenue**. Do not take a simple average of the row-level profit-margin percentages.

For order counts, count distinct `Order Number` values rather than transaction lines, because an order can contain multiple line items.

## 5. Validation Performed

The uploaded workbook was inspected read-only; no formulas, data types, query connections, PivotTables, or source records were changed.

Checks performed:
- All 62,884 Sales rows were present.
- All 10 formula columns contained formulas throughout the Sales data rows.
- No `Not Found` results were detected in the cached results of those formula columns.
- Cached `Total Revenue` values matched `Quantity × Unit Price`.
- Cached `Gross Profit` values matched `Total Revenue − Total Cost`.
- Cached `Profit Margin` values matched `Gross Profit ÷ Total Revenue` (for non-zero revenue rows).
- All `StoreKey` values matched a record in the Stores table.
- The stored `Store State` values matched the corresponding Stores-table state values.

## 6. Compatibility and Maintenance

- `XLOOKUP` requires Microsoft 365 or Excel 2021 and later.
- The source workbook includes Power Query connections, PivotTables, and an Excel Data Model. Avoid round-tripping it through tools that do not fully preserve these Excel features.
- If source tables are refreshed or replaced, refresh the Power Query outputs and Data Model in Excel, then check PivotTables and calculated columns.
