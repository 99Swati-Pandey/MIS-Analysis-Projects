# Sales Data Cleaning: Excel Power Query

## Overview
A messy e-commerce sales dataset (335 raw rows) cleaned end to end with **Power Query in Excel**. The cleaning lives in one repeatable query, so new or corrected raw data can be cleaned again with a single **Refresh All**.

**Tool:** Microsoft Excel (Power Query / M language)

## Workbook structure
| Sheet | Purpose |
|---|---|
| `Sales_Raw` | Original untouched data (Excel table `Table1`), the query source |
| `Sales_Clean` | Output of the `Sales_Clean` Power Query (236 rows, 13 columns) |

## Data quality issues found and how I fixed them
| Issue | Example | Fix |
|---|---|---|
| Exact duplicate orders | 15 repeated rows | Removed duplicates on `OrderID` |
| Invalid quantity | Quantity = 0 or -1 (84 rows) | Filtered out (`Quantity > 0`) |
| Mixed date formats | `2024-05-04`, `04-05-2024`, `05/04/2024`, `4 May 2024` | Custom M function parses each format to a real date |
| Ambiguous dates | `02/04/2024` (day/month or month/day?) | Assumed day/month; flagged in `DateAmbiguous` |
| Inconsistent IDs | `C6`, `c018`, `C-36`, `P-2`, ` P015 ` | Standardised to `C006`, `C018`, `C036`, `P002`, `P015` |
| Messy text | extra/inner spaces, `reyansh joshi`, `SNEAKERS` | Trim + collapse double spaces + proper case |
| Duplicate region names | Madras, Bombay, Bangalore, New Delhi, Calcutta | Mapped to Chennai, Mumbai, Bengaluru, Delhi, Kolkata |
| Payment labels | `COD`, `Cash on Delivery`, `upi` | Unified to `Cash on Delivery` and `UPI` |
| Status casing | `delivered` vs `Delivered` | Proper case |
| Currency text in price | `₹2,558.72`, `Rs. 929`, ` 1297.16 ` | Stripped symbols, commas and spaces; converted to number |
| Missing discount | 67 blanks | Set to 0 (assumes blank = no discount) |

## Output columns (`Sales_Clean`)
`OrderID`, `OrderDate`, `DateAmbiguous`, `CustomerID`, `CustomerName`, `ProductID`, `Region`, `ProductName`, `Quantity`, `UnitPrice`, `Discount`, `PaymentMethod`, `Status`

## Validation checks
- `OrderID` is unique (236 orders)
- Each `CustomerID` maps to exactly one customer name, and each `ProductID` to exactly one product name
- No blank dates or IDs; no quantity ≤ 0
- Prices, quantities and discounts match the raw values after cleaning

## Assumptions and limitations
- **45 dates** with a `/` separator and both parts ≤ 12 cannot be resolved from the data alone. Filter `DateAmbiguous = Yes` to review them.
- **Blank discounts** are treated as 0.
- **Cancelled and Returned orders** are kept in the data. Filter `Status = Delivered` for realised revenue.
- `Kaithal` is a town rather than a major city like the other regions. It is kept as in the source.

## How to use
1. Open the workbook in Excel.
2. Go to **Data → Refresh All** to re-run the cleaning query.
3. To review the steps, open **Data → Queries & Connections → Sales_Clean → Edit** (Power Query Editor).

## Skills demonstrated
Power Query, M functions, data profiling, text and date standardisation, deduplication, data validation, documentation
