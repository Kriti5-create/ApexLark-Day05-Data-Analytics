# ApexLark Data Analytics Internship — Day 05
## Advanced Business Analysis Using Excel Lookup Functions & Dynamic Reporting

### Deliverables
- `ApexLark_Day05_Advanced_Business_Analysis.xlsx` — integrated workbook with source tables, lookup formulas, logical classifications, KPI summary, charts and transaction-level report.
- `ApexLark_Day05_Dynamic_Business_Report.pdf` — business performance findings, data limitations and recommendations.
- `README_Day05.md` — project documentation.

### Workbook structure
| Worksheet | Purpose |
|---|---|
| Final_Report | KPI summary, regional/category/customer/product summaries, charts, conditional formatting and transaction-level lookup output |
| Sales_Data | 9,993 cleaned sales transaction rows |
| Customer_Data | 793 unique customer IDs, segment and customer-level sales/profit metrics |
| Product_Data | Product master with category, sub-category, average realized unit price and product-level metrics |
| Lookup_Functions | Examples/documentation for XLOOKUP, VLOOKUP, HLOOKUP, INDEX+MATCH, IF, IFERROR, AND and OR |

### Key results
- Total Sales: $2,296,919.49
- Total Profit: $286,409.08
- Total Orders: 5,009
- Total Customers: 793
- Profit Margin: 12.47%
- Highest revenue region: West ($725,457.82)
- Highest-sales customer ID: SM-20320 ($25,043.05)
- Most profitable category: Technology ($145,454.95 profit)

### Formula notes
- XLOOKUP retrieves customer and product details.
- VLOOKUP retrieves product category.
- INDEX + MATCH retrieves average realized unit price.
- HLOOKUP is demonstrated in the Lookup_Functions worksheet.
- IF categorizes sales; IFERROR handles missing lookup results; AND identifies profitable transactions; OR flags high-discount or loss-making transactions.

### Data limitations
The supplied source does not contain customer names or list prices. Customer IDs are used for customer-level analysis. Product Unit Price is calculated as average realized Sales divided by Quantity, and should not be interpreted as list price. These limitations are disclosed in the PDF report and workbook.

### How to use
Open the `.xlsx` file in Microsoft Excel and allow formulas to recalculate. The transaction-level lookup table and summary formulas are designed to update when source data is edited within the referenced ranges. For a refreshed dataset, replace the data rows while preserving the worksheet column structure and refresh/recalculate the workbook.
