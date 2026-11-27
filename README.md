# Apple Inc. Financial Statement Extraction from SEC EDGAR (FY2021–FY2026)

Beginner-friendly Python scripts that pull Apple's historical financial statement data from the official **SEC EDGAR Company Facts API**, clean and validate it, and load it into the `Raw IS`, `Raw BS` and `Raw CF` sheets of an Excel financial model. The data feeds a three-statement model and DCF valuation of Apple. The forecasting and valuation work itself is not part of this repository.

- **Source:** SEC EDGAR only (`data.sec.gov`). No third-party financial data providers are used for extraction.
- **Scope:** annual figures from Apple's **10-K filings**, fiscal years 2021–2026.
- **Output:** Income Statement, Balance Sheet and Cash Flow line items by fiscal year, in USD millions.

> **Disclaimer:** This is an educational project, not financial advice, and it is not affiliated with Apple Inc. or the SEC. The scripts reconcile the data internally, but you should still spot-check key figures against the 10-K filings before relying on them.

---

## How the SEC Company Facts API works

The SEC publishes every number a company tags in its XBRL filings through a single free endpoint:

```
https://data.sec.gov/api/xbrl/companyfacts/CIK0000320193.json
```

- `0000320193` is Apple's CIK (SEC company ID), padded to 10 digits.
- One request returns every tagged fact the company has ever reported, as a large JSON file. Data is nested as `facts → us-gaap → <tag> → units → USD → [data points]`.
- The SEC requires a **descriptive `User-Agent` header** (your name and email). Requests without one are blocked. The fair-access limit is 10 requests per second; this project makes a single request.
- The same number appears many times (for example, a FY2023 figure is repeated as a comparative in later 10-Ks), and the file mixes 10-Q and 10-K data. The scripts filter and de-duplicate it.

---

## Repository contents

| File | Purpose |
|---|---|
| `step1_download.py` | Downloads the Company Facts JSON and saves it as `apple_companyfacts.json`. The only script that contacts the SEC. |
| `sec_helpers.py` | Shared helper functions. Loads the saved JSON and returns clean annual values for any tag. |
| `step5_build.py` | Builds the Income Statement, Balance Sheet and Cash Flow tables (FY2021–FY2025), runs consistency checks, and exports them to the `Raw IS`, `Raw BS` and `Raw CF` sheets. |
| `step6_coverage.py` | Optional tag-discovery tool. Reports which of the additional line items exist for each year, and searches every namespace for tags containing keywords. Useful if you adapt this to another company. |
| `step7_extract_and_append.py` | Extracts the additional line items, validates them against reported totals, and appends them as new rows to `Raw BS` and `Raw CF`. |
| `step8_add_fy2026.py` | Adds an `FY2026` column to all three Raw sheets and adds the `Intangible assets, net` row. |

The scripts are numbered by development order. Steps 2–4 were early tests and are not included.

---

## Requirements

- Python 3.11 (developed with Windows, VS Code and an Anaconda environment, but not OS-specific)
- `pandas`, `requests`, `openpyxl`

```
pip install pandas requests openpyxl
```

or, with conda: `conda install pandas requests openpyxl`

---

## Setup

1. **Set your User-Agent.** In `step1_download.py`, replace the placeholder with your own name and a real email address. The SEC uses this to contact you if your requests cause problems:
   ```python
   HEADERS = {"User-Agent": "Your Name your.email@example.com"}
   ```
2. **Create the Excel workbook.** The scripts write into an existing workbook named `Apple_Financial_Model.xlsx` in the same folder. Create a blank workbook with that name (it needs at least one sheet). If the file is missing, the scripts stop and write nothing.
3. **Close the workbook in Excel** before every run, otherwise saving fails with a permission error.

---

## Run order

```
python step1_download.py              # 1. download the SEC data (re-run to refresh it)
python step5_build.py                 # 2. build tables, run checks, create Raw IS / Raw BS / Raw CF
python step6_coverage.py              # (optional) tag coverage and discovery
python step7_extract_and_append.py    # 3. append the additional line items to Raw BS and Raw CF
python step8_add_fy2026.py            # 4. add the FY2026 column and Intangible assets row
```

`apple_companyfacts.json` is a snapshot. Re-run `step1_download.py` to pick up new filings, such as a new 10-K.

---

## What each script does

### `step1_download.py`
Sends one request to the SEC, saves the response to `apple_companyfacts.json`, and prints basic validation (company name, CIK, number of `us-gaap` tags, and an example data point). It has error handling for network and HTTP failures.

### `sec_helpers.py`
Loads the saved JSON once and provides:

- `get_annual_duration(tag, ...)` for income statement and cash flow items, which cover a period (they have a start and end date).
- `get_annual_instant(tag, ...)` for balance sheet items, which are year-end snapshots (end date only).
- `get_annual_multi(candidates, ...)` tries several candidate tags in priority order, uses the first with a value for each fiscal year, and reports which tag supplied it.

### `step5_build.py`
Extracts the core line items, builds one table per statement with fiscal years as columns, and runs six consistency checks (see Validation). It then exports to the workbook in USD millions.

### `step7_extract_and_append.py`
Extracts additional items needed for an integrated three-statement model, validates them against reported totals, and appends them under the existing rows of `Raw BS` and `Raw CF`. Each row has the SEC tag and any sign change recorded in a notes column. A sheet is written only if its checks pass.

### `step8_add_fy2026.py`
Adds the `FY2026` column next to `FY2025` in each Raw sheet and fills every existing row. It also adds the `Intangible assets, net` row to `Raw BS`, which Apple began reporting separately from FY2025 onwards (see Intangible assets under mapping decisions).

---

## Methodology

- **10-K only.** Quarterly (10-Q) data is excluded.
- **Fiscal year comes from the period end date.** Apple's fiscal year ends on the last Saturday of September, so the calendar year of the end date is the fiscal year. The API's own `fy` field is the fiscal year of the *filing*, not of the period, so it isn't used.
- **Full-year periods only.** Duration items are kept only if the period is 350–380 days. The wide range is deliberate: Apple uses a 52/53-week year, and FY2023 was a 53-week year (371 days).
- **De-duplication.** When the same period appears in several 10-Ks, the **most recently filed** value is kept, so restated figures win.
- **Units.** Monetary values are in USD millions. Diluted EPS is in USD per share, and diluted shares are in millions.
- **No overwriting.** Existing cells and rows are never overwritten. Labels already in a sheet are skipped, so re-running a script is safe.

---

## Line-item mapping

SEC (`us-gaap`) tags don't always match the names on Apple's statements. The mapping below was verified with coverage checks and reconciliations.

### Income statement (`Raw IS`)
| Line item | SEC tag |
|---|---|
| Net sales | `RevenueFromContractWithCustomerExcludingAssessedTax` |
| Cost of sales | `CostOfGoodsAndServicesSold` |
| Gross profit | `GrossProfit` |
| R&D | `ResearchAndDevelopmentExpense` |
| SG&A | `SellingGeneralAndAdministrativeExpense` |
| Total operating expenses | `OperatingExpenses` |
| Operating income | `OperatingIncomeLoss` |
| Other income/(expense), net | `NonoperatingIncomeExpense` |
| Income before taxes | `IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest` |
| Income tax expense | `IncomeTaxExpenseBenefit` |
| Net income | `NetIncomeLoss` |
| Diluted EPS | `EarningsPerShareDiluted` (unit `USD/shares`) |
| Diluted shares | `WeightedAverageNumberOfDilutedSharesOutstanding` (unit `shares`) |

### Balance sheet (`Raw BS`)
| Line item | SEC tag |
|---|---|
| Cash and equivalents | `CashAndCashEquivalentsAtCarryingValue` |
| Current marketable securities | `MarketableSecuritiesCurrent` |
| Accounts receivable | `AccountsReceivableNetCurrent` |
| Vendor non-trade receivables | `NontradeReceivablesCurrent` |
| Inventories | `InventoryNet` |
| Total current assets | `AssetsCurrent` |
| Non-current marketable securities | `MarketableSecuritiesNoncurrent` |
| PP&E, net | `PropertyPlantAndEquipmentNet` |
| Total assets | `Assets` |
| Accounts payable | `AccountsPayableCurrent` |
| Commercial paper | `CommercialPaper` |
| Current term debt | `LongTermDebtCurrent` |
| Total current liabilities | `LiabilitiesCurrent` |
| Non-current term debt | `LongTermDebtNoncurrent` |
| Total liabilities | `Liabilities` |
| Shareholders' equity | `StockholdersEquity` |
| Other current assets | `OtherAssetsCurrent` |
| Other non-current assets | `OtherAssetsNoncurrent` |
| Deferred revenue (current) | `ContractWithCustomerLiabilityCurrent` |
| Other current liabilities | `OtherLiabilitiesCurrent` |
| Other non-current liabilities | `OtherLiabilitiesNoncurrent` |
| Intangible assets, net | See "Intangible assets" below |

### Cash flow (`Raw CF`)
| Line item | SEC tag | Sign |
|---|---|---|
| Net income | `NetIncomeLoss` | as reported |
| D&A | `DepreciationDepletionAndAmortization` | as reported |
| Share-based compensation | `ShareBasedCompensation` | as reported |
| Operating cash flow | `NetCashProvidedByUsedInOperatingActivities` | as reported |
| Capex | `PaymentsToAcquirePropertyPlantAndEquipment` | outflow shown **positive** |
| Investing cash flow | `NetCashProvidedByUsedInInvestingActivities` | as reported |
| Dividends paid | `PaymentsOfDividends` | outflow shown **positive** |
| Share repurchases | `PaymentsForRepurchaseOfCommonStock` | outflow shown **positive** |
| Financing cash flow | `NetCashProvidedByUsedInFinancingActivities` | as reported |
| Other non-cash adjustments (incl. deferred income tax) | `OtherNoncashIncomeExpense` | flipped |
| Chg in accounts receivable | `IncreaseDecreaseInAccountsReceivable` | flipped |
| Chg in vendor non-trade receivables | `IncreaseDecreaseInOtherReceivables` | flipped |
| Chg in inventories | `IncreaseDecreaseInInventories` | flipped |
| Chg in other current & non-current assets | `IncreaseDecreaseInOtherOperatingAssets` | flipped |
| Chg in accounts payable | `IncreaseDecreaseInAccountsPayable` | kept |
| Chg in other current & non-current liabilities | `IncreaseDecreaseInOtherOperatingLiabilities` | kept |
| Purchases of marketable securities | `PaymentsToAcquireAvailableForSaleSecuritiesDebt` | flipped |
| Proceeds from maturities of marketable securities | `ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities` | kept |
| Proceeds from sales of marketable securities | `ProceedsFromSaleOfAvailableForSaleSecuritiesDebt` | kept |
| Other investing activities | `PaymentsForProceedsFromOtherInvestingActivities` | flipped |
| Proceeds from issuance of debt | `ProceedsFromIssuanceOfLongTermDebt` | kept |
| Repayments of debt | `RepaymentsOfLongTermDebt` | flipped |
| Commercial paper, net | `ProceedsFromRepaymentsOfCommercialPaper` | kept |
| Taxes paid on net share settlement | `PaymentsRelatedToTaxWithholdingForShareBasedCompensation` | flipped |
| Other financing activities | `ProceedsFromPaymentsForOtherFinancingActivities` | kept |
| Net change in cash | `CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect` | kept |
| Cash, cash equivalents & restricted cash, end of year | `CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents` | kept |
| Cash, cash equivalents & restricted cash, beginning of year | same tag, prior year-end value | kept |

### Sign conventions
- **Core cash flow rows** (Capex, Dividends paid, Share repurchases) keep the SEC's positive values, which are outflows.
- **Additional cash flow rows** follow the **printed cash flow statement**: outflows are negative. XBRL reports payments and increases in assets as positive, so those rows are multiplied by -1. Free cash flow in a model is therefore operating cash flow minus Capex, using the positive Capex row.
- Treat the two conventions consistently when linking a model.
- The sign of "Other non-cash adjustments" is detected by the script (it tests which sign makes operating cash flow reconcile) and recorded in the notes column.

---

## Mapping decisions and data caveats

- **Revenue:** Apple's total revenue is only available under `RevenueFromContractWithCustomerExcludingAssessedTax` for all years. `Revenues` and `SalesRevenueNet` return nothing in the annual data.
- **Deferred income tax and acquisitions are not separate rows.** Apple's FY2023 10-K merged them into "Other" (operating) and "Other investing activities". Because the latest filing wins for each year, FY2021–22 use the merged presentation, and adding separate rows would double count (the reconciliation checks fail if they are added).
- **Non-current deferred revenue is not available.** The tag has no annual 10-K values after FY2018. It sits inside "Other non-current liabilities". Do not add it separately.
- **FX effect on cash is not available.** Apple's cash flow statement has no such line.
- **Vendor non-trade receivables (cash flow)** uses the generic tag `IncreaseDecreaseInOtherReceivables`, mapped by name rather than by an Apple-specific tag.
- **"Other" operating assets and liabilities** are single combined lines. SEC tags don't split current from non-current.
- **Net change in cash includes restricted cash**, so it can differ slightly from the change in balance sheet cash.
- **Interest expense is not extracted.** Apple disclosed it in the notes for FY2021–23, but it could not be found for FY2024–26. Only the net `Other income/(expense), net` is available for those years.

### Intangible assets (`step8_add_fy2026.py`)
From FY2025 onwards, Apple's balance sheet shows `Intangible assets, net` separately. In earlier years it was inside "Other non-current assets".

- The script looks for a reported tag whose FY2025 value equals 11,093 (USD millions), searching every namespace. If it finds none, it **derives** the figure as total assets less current assets, non-current marketable securities, PP&E and other non-current assets, and proceeds only if FY2025 matches 11,093. A derived row is labelled "Derived: ..." in the notes column.
- FY2021–FY2024 are left blank, since the line wasn't separate.
- **Series break:** "Other non-current assets" includes intangibles in FY2021–24 but excludes them from FY2025 onwards.
- The FY2025 "Other non-current assets" value from the original extraction is the old combined figure. The `RECAST_FY2025_OTHER_NCA` flag in `step8_add_fy2026.py` (default `False`) updates that one cell to the recast value. The script prints both values for comparison either way.

---

## Validation

`step5_build.py` checks (within a $1m rounding tolerance):

1. Net sales − Cost of sales = Gross profit
2. Gross profit − Operating expenses = Operating income
3. Operating income + Other income = Income before taxes
4. Income before taxes − Tax = Net income
5. Total assets = Total liabilities + Shareholders' equity
6. Income statement net income = Cash flow net income

`step7_extract_and_append.py` checks:

- **Balance sheet:** current assets, total assets, current liabilities and total liabilities each equal the sum of their components.
- **Operating cash flow** reconciles, and "Other non-cash" is assigned the sign that makes it tie.
- **Investing and financing** components sum to the reported totals.
- **Operating + investing + financing** = net change in cash.
- **Beginning cash + net change = ending cash.**

If a sheet's checks fail, nothing is written to that sheet.

---

## Safety features

- Existing sheets, rows and cells are never overwritten. Existing labels are skipped.
- A backup copy of the workbook is saved before each write (`Apple_Financial_Model_backup.xlsx`, and `Apple_Financial_Model_backup_before_FY2026.xlsx` for step 8).
- Sheet-name clash checks and "row not empty" checks stop a write rather than overwrite data.
- A `PermissionError` (usually the workbook being open in Excel) produces a clear message.
- Re-running any script is safe: it adds only what's missing.

---

## Workbook layout

Each Raw sheet has labels in column A and fiscal years across the top (`FY2021` … `FY2026`). `Raw BS` and `Raw CF` have a notes column after the last year: `SEC tag (sign applied)`. When `step8_add_fy2026.py` adds the `FY2026` column, it **moves that notes column one place to the right**. Row 1 of `Raw IS` has no notes column.

---

## Known limitations

1. **Apple-specific.** Tag mappings, sign conventions, the intangible assets check figure (11,093) and the fiscal calendar assumptions are specific to Apple. Adapting this to another company requires re-running the coverage tool (`step6_coverage.py`) and re-checking each mapping.
2. **Restatements.** The latest filing wins for each period. If Apple recasts its presentation in a new 10-K, as it did in FY2023 and for intangibles in FY2026, the newest column uses the new presentation while older columns keep the earlier one. Compare overlapping years after adding a new filing.
3. **Step 7's checks target the FY2025-era presentation.** Its balance sheet identities were written for data as presented in the FY2025 10-K. If you run it against a JSON that already includes the FY2026 10-K, the FY2025 total assets check will likely fail, because FY2025 "Other non-current assets" is now recast without intangibles. If this happens, run steps 1, 5 and 7 as described above using a data snapshot taken before the FY2026 10-K (or adjust that check), then run step 8.
4. **`openpyxl` limitations.** Saving a workbook with `openpyxl` can drop charts, images, pivot tables and some conditional formatting. Check your workbook against the backup after the first run.
5. **Snapshot data.** The JSON is a local copy. Re-run `step1_download.py` for new filings.
6. **Reconciliation is not verification.** Passing checks show the data is internally consistent, not that every figure matches the filing. Spot-check against Apple's 10-Ks.

---

## Data source and fair use

All data comes from SEC EDGAR via `data.sec.gov`. Use a real `User-Agent`, stay within the SEC's 10 requests-per-second limit, and see the SEC's [fair access guidance](https://www.sec.gov/os/webmaster-faq#developers).
