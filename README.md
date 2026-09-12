# Excel WDV Depreciation UDF (Date Range & Financial Year Aware)

A custom VBA User-Defined Function (UDF) for Microsoft Excel that calculates **Written Down Value (WDV) Depreciation** for any custom date range across an **April–March Financial Year (FY)**.

This repository provides a clean, maintainable VBA alternative to long, unreadable cell formulas or `LAMBDA` functions, making it compatible with legacy Excel versions.

---

## Features

- **April–March FY Support:** Automatically aligns asset purchase dates and calculation periods with Indian / standard April–March financial years.
- **Pro-Rata Calculations:** Accurately calculates partial-year depreciation for Year 1 (based on purchase date) and partial periods for target date ranges.
- **Salvage Value Floor:** Ensures asset WDV never drops below specified salvage values.
- **No LAMBDA Required:** Fully compatible with older versions of Excel where `LAMBDA` and dynamic arrays are unavailable.
- **Form-Fitted Parameters:** Easily invoke it as a native worksheet function with clear inputs.

---

## Installation

1. Open your Excel workbook and press `Alt + F11` to launch the **VBA Editor**.
2. Click **Insert > Module** from the top menu.
3. Copy the code from [`WDV_Depreciation.bas`](./WDV_Depreciation.bas) (or `Module1.vba`) and paste it into the code window.
4. Save your workbook as an **Excel Macro-Enabled Workbook (`.xlsm`)**.
5. Copy the code from ['Helper Description.cls'] and paste it in ThisWorkbook code window.


---

## Usage

Use the function directly in any cell just like a standard Excel formula:

```excel
=WDV_Depreciation(Cost, Rate, Salvage, Life, PurchaseDate, RangeStart, RangeEnd)
