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
- **Function Argument Tooltips:** Registers itself with Excel's Function Wizard, so argument names and descriptions show up automatically when typing the formula.

---

## Releases

This project is distributed in two ways — pick whichever suits you:

| Release | Contents | Best for |
|---|---|---|
| **[V1 — Source Files](https://github.com/rocks1997/Excel-VBA-for-WDV-Depriciation-calculation/releases/tag/V1)** | Raw `.bas` and `.cls` files | Developers who want to inspect the code, modify it, or add it to an existing project manually |
| **[V2 — Self-Installing Add-in](https://github.com/rocks1997/Excel-VBA-for-WDV-Depriciation-calculation/releases/tag/V2)** | `WDV_Depreciation_Installer.xlsm` | Everyone else — one file, one click, no VBA Editor required |

### Recommended: V2, the self-installing workbook

1. Download `WDV_Depreciation_Installer.xlsm` from the [V2 release](https://github.com/rocks1997/Excel-VBA-for-WDV-Depriciation-calculation/releases/tag/V2).
2. Open it in Excel and click **Enable Content** on the security bar (required — it contains the installer macro).
3. Click **Yes** on the prompt that appears, or use the **Install Add-in** button on the sheet.
4. Done — `WDV_Depreciation()` is now registered as a proper Excel add-in and will be available in every workbook, every time you open Excel.
5. To remove it later, open the same `.xlsm` again and run the `UninstallAddin` macro (`Alt+F8` → `UninstallAddin` → `Run`).

### Manual install: V1, raw source

Use this if you want the code itself — to read it, adapt it, or wire it into your own add-in.

1. Open your Excel workbook and press `Alt + F11` to launch the **VBA Editor**.
2. Click **Insert > Module** from the top menu.
3. Copy the code from [`WDV.Depriciation.bas`](https://github.com/rocks1997/Excel-VBA-for-WDV-Depriciation-calculation/releases/download/V1/WDV.Depriciation.bas) and paste it into the code window.
4. Save your workbook as an **Excel Macro-Enabled Workbook (`.xlsm`)**.
5. Copy the code from [`Helper.Description.cls`](https://github.com/rocks1997/Excel-VBA-for-WDV-Depriciation-calculation/releases/download/V1/Helper.Description.cls) and paste it into the **ThisWorkbook** code window — this registers the function's argument descriptions with Excel's Function Wizard.

---

## Usage

Use the function directly in any cell just like a standard Excel formula:

```excel
=WDV_Depreciation(Cost, Rate, Salvage, Life, PurchaseDate, RangeStart, RangeEnd)
```

| Argument | Description |
|---|---|
| `Cost` | Original purchase cost of the asset |
| `Rate` | WDV depreciation rate, as a decimal (e.g. `0.2209` for 22.09%) |
| `Salvage` | Residual/salvage value — depreciation stops once WDV reaches this floor |
| `Life` | Useful life of the asset, in years |
| `PurchaseDate` | The date the asset was acquired |
| `RangeStart` | Start date of the period you want depreciation for |
| `RangeEnd` | End date of the period you want depreciation for |

**Example:**

```excel
=WDV_Depreciation(12000, 0.2209, 600, 12, DATE(2023,4,1), DATE(2024,4,1), DATE(2025,3,31))
```

Returns the depreciation for the second financial year (1-Apr-2024 to 31-Mar-2025) of an asset costing ₹12,000, bought on 1-Apr-2023, with a 22.09% WDV rate, ₹600 salvage value, and a 12-year useful life.

The result is returned **unrounded** — wrap it in Excel's own `ROUND()` where you need a whole-rupee figure:

```excel
=ROUND(WDV_Depreciation(...), 0)
```

This also means you can call the function for several shorter sub-periods (e.g. month by month) and sum them without compounding rounding errors, then round only the final total.
