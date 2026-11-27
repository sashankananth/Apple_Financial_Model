import os
import shutil
from copy import copy
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from sec_helpers import get_annual_multi, data

MODEL_FILE = "Apple_Financial_Model.xlsx"
YEAR = 2026
HEADER = f"FY{YEAR}"

CASH_TAG = "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"
BEGIN = "Cash, cash equivalents & restricted cash - beginning of year"
END = "Cash, cash equivalents & restricted cash - end of year"

# Units that are NOT USD millions: EPS stays in USD, shares are shown in millions
UNITS = {"Diluted EPS": "USD/shares", "Diluted shares": "shares"}

# --- Line items per sheet: (label as it appears in column A, SEC tag(s), sign) ---
# sign = -1 flips the SEC value so it matches the printed cash flow statement.
ITEMS = {
    "Raw IS": [
        ("Net sales", "RevenueFromContractWithCustomerExcludingAssessedTax", 1),
        ("Cost of sales", "CostOfGoodsAndServicesSold", 1),
        ("Gross profit", "GrossProfit", 1),
        ("R&D", "ResearchAndDevelopmentExpense", 1),
        ("SG&A", "SellingGeneralAndAdministrativeExpense", 1),
        ("Total operating expenses", "OperatingExpenses", 1),
        ("Operating income", "OperatingIncomeLoss", 1),
        ("Other income/(expense)", "NonoperatingIncomeExpense", 1),
        ("Income before taxes",
         "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest", 1),
        ("Income tax expense", "IncomeTaxExpenseBenefit", 1),
        ("Net income", "NetIncomeLoss", 1),
        ("Diluted EPS", "EarningsPerShareDiluted", 1),
        ("Diluted shares", "WeightedAverageNumberOfDilutedSharesOutstanding", 1),
    ],
    "Raw BS": [
        ("Cash and equivalents", "CashAndCashEquivalentsAtCarryingValue", 1),
        ("Current marketable securities", "MarketableSecuritiesCurrent", 1),
        ("Accounts receivable", "AccountsReceivableNetCurrent", 1),
        ("Vendor non-trade receivables", "NontradeReceivablesCurrent", 1),
        ("Inventories", "InventoryNet", 1),
        ("Total current assets", "AssetsCurrent", 1),
        ("Non-current marketable securities", "MarketableSecuritiesNoncurrent", 1),
        ("PP&E, net", "PropertyPlantAndEquipmentNet", 1),
        ("Total assets", "Assets", 1),
        ("Accounts payable", "AccountsPayableCurrent", 1),
        ("Commercial paper", "CommercialPaper", 1),
        ("Current term debt", "LongTermDebtCurrent", 1),
        ("Total current liabilities", "LiabilitiesCurrent", 1),
        ("Non-current term debt", "LongTermDebtNoncurrent", 1),
        ("Total liabilities", "Liabilities", 1),
        ("Shareholders' equity", "StockholdersEquity", 1),
        ("Other current assets", "OtherAssetsCurrent", 1),
        ("Other non-current assets", "OtherAssetsNoncurrent", 1),
        ("Deferred revenue (current)",
         ["ContractWithCustomerLiabilityCurrent", "DeferredRevenueCurrent"], 1),
        ("Other current liabilities", "OtherLiabilitiesCurrent", 1),
        ("Other non-current liabilities", "OtherLiabilitiesNoncurrent", 1),
    ],
    "Raw CF": [
        # Original rows (signs as originally extracted: outflows positive)
        ("Net income", "NetIncomeLoss", 1),
        ("D&A", "DepreciationDepletionAndAmortization", 1),
        ("Share-based compensation", "ShareBasedCompensation", 1),
        ("Operating cash flow", "NetCashProvidedByUsedInOperatingActivities", 1),
        ("Capex", "PaymentsToAcquirePropertyPlantAndEquipment", 1),
        ("Investing cash flow", "NetCashProvidedByUsedInInvestingActivities", 1),
        ("Dividends paid", "PaymentsOfDividends", 1),
        ("Share repurchases", "PaymentsForRepurchaseOfCommonStock", 1),
        ("Financing cash flow", "NetCashProvidedByUsedInFinancingActivities", 1),
        # Additional rows (signs as printed on the cash flow statement)
        ("Other non-cash adjustments (incl. deferred income tax)", "OtherNoncashIncomeExpense", -1),
        ("Chg in accounts receivable", "IncreaseDecreaseInAccountsReceivable", -1),
        ("Chg in vendor non-trade receivables", "IncreaseDecreaseInOtherReceivables", -1),
        ("Chg in inventories", "IncreaseDecreaseInInventories", -1),
        ("Chg in other current & non-current assets", "IncreaseDecreaseInOtherOperatingAssets", -1),
        ("Chg in accounts payable", "IncreaseDecreaseInAccountsPayable", 1),
        ("Chg in other current & non-current liabilities", "IncreaseDecreaseInOtherOperatingLiabilities", 1),
        ("Purchases of marketable securities",
         ["PaymentsToAcquireAvailableForSaleSecuritiesDebt", "PaymentsToAcquireAvailableForSaleSecurities"], -1),
        ("Proceeds from maturities of marketable securities",
         ["ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities",
          "ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecuritiesDebt"], 1),
        ("Proceeds from sales of marketable securities",
         ["ProceedsFromSaleOfAvailableForSaleSecuritiesDebt", "ProceedsFromSaleOfAvailableForSaleSecurities"], 1),
        ("Other investing activities", "PaymentsForProceedsFromOtherInvestingActivities", -1),
        ("Proceeds from issuance of debt", "ProceedsFromIssuanceOfLongTermDebt", 1),
        ("Repayments of debt", "RepaymentsOfLongTermDebt", -1),
        ("Commercial paper, net", "ProceedsFromRepaymentsOfCommercialPaper", 1),
        ("Taxes paid on net share settlement",
         "PaymentsRelatedToTaxWithholdingForShareBasedCompensation", -1),
        ("Other financing activities", "ProceedsFromPaymentsForOtherFinancingActivities", 1),
        ("Net change in cash",
         ["CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseExcludingExchangeRateEffect",
          "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect",
          "CashAndCashEquivalentsPeriodIncreaseDecrease"], 1),
        (BEGIN, CASH_TAG, 1),
        (END, CASH_TAG, 1),
    ],
}


def fetch(sheet, label, tags, sign):
    """Get the FY2026 value for one line item, in the units used in the sheet.
    Returns (value, tag_used), or (None, None) if the SEC data has no value."""
    tags = [tags] if isinstance(tags, str) else tags
    unit = UNITS.get(label, "USD")
    kind = "instant" if sheet == "Raw BS" or label in (BEGIN, END) else "duration"
    # Beginning cash = the FY2025 year-end balance; everything else is FY2026
    year = YEAR - 1 if label == BEGIN else YEAR
    result = get_annual_multi(tags, kind=kind, unit=unit, first_fy=year, last_fy=year)
    if result is None or result.empty:
        return None, None
    divisor = 1 if unit == "USD/shares" else 1e6     # USD -> millions, shares -> millions
    value = float(result["val"].iloc[0]) * sign / divisor
    return value, result["tag"].iloc[0]


def add_year_column(ws):
    """Find the FY2026 column, or create it right after FY2025. Returns its column number."""
    headers = {str(c.value): c.column for c in ws[1] if c.value is not None}
    if HEADER in headers:
        return headers[HEADER]                        # already there (safe to re-run)
    if "FY2025" not in headers:
        print(f"  {ws.title}: no FY2025 header in row 1. Skipped.")
        return None

    col = headers["FY2025"] + 1
    last_row = ws.max_row
    # If the next column is taken (the "SEC tag" notes column), move it one column right
    if any(ws.cell(r, col).value is not None for r in range(1, last_row + 1)):
        if any(ws.cell(r, col + 1).value is not None for r in range(1, last_row + 1)):
            print(f"  {ws.title}: no free column next to FY2025. Skipped.")
            return None
        letter = get_column_letter(col)
        ws.move_range(f"{letter}1:{letter}{last_row}", cols=1)

    ws.cell(1, col, HEADER)
    ws.cell(1, col)._style = copy(ws.cell(1, col - 1)._style)   # same look as the FY2025 header
    width = ws.column_dimensions[get_column_letter(col - 1)].width
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width
    return col


def fill_sheet(ws, items):
    """Write one FY2026 value next to each matching label in column A."""
    col = add_year_column(ws)
    if col is None:
        return 0
    wanted = {label: (tags, sign) for label, tags, sign in items}
    written = 0
    for row in range(2, ws.max_row + 1):
        raw_label = ws.cell(row, 1).value
        if raw_label is None:
            continue
        label = str(raw_label).strip()
        if label == INTANGIBLE_LABEL:
            continue                                  # handled separately in add_intangibles()
        if label not in wanted:
            print(f"  {ws.title}: no mapping for '{label}' - left blank")
            continue
        if ws.cell(row, col).value is not None:
            continue                                  # never overwrite an existing value
        tags, sign = wanted[label]
        value, tag_used = fetch(ws.title, label, tags, sign)
        if value is None:
            print(f"  {ws.title}: no {HEADER} value found for '{label}'")
            continue
        cell = ws.cell(row, col, value)
        cell._style = copy(ws.cell(row, col - 1)._style)          # same number format as FY2025
        first_tag = tags if isinstance(tags, str) else tags[0]
        if tag_used != first_tag:
            print(f"  NOTE: '{label}' used fallback tag {tag_used}")
        written += 1
    return written

# --- Intangible assets (a separate line from FY2025 onwards) --------------------
INTANGIBLE_LABEL = "Intangible assets, net"
INTANGIBLE_FY2025_CHECK = 11093       # FY2025 balance in USD millions, per Apple's FY2026 filings
MISSING_MARK = None                   # years with no data stay blank; use "-" to show a dash
RECAST_FY2025_OTHER_NCA = False       # True = update FY2025 'Other non-current assets' to the recast figure


def find_reported_intangible():
    """Look in every namespace (us-gaap, aapl, ...) for a year-end tag whose FY2025 value
    matches the check figure. Returns (tag, {year: value in USD}) or None."""
    for namespace, tags in data["facts"].items():
        for tag, info in tags.items():
            if "intangible" not in tag.lower():
                continue
            values = {}
            for p in sorted(info.get("units", {}).get("USD", []), key=lambda p: p["filed"]):
                if p["form"] == "10-K" and "start" not in p:      # year-end balances only
                    values[int(p["end"][:4])] = p["val"]          # later filings overwrite earlier ones
            if round(values.get(2025, 0) / 1e6) == INTANGIBLE_FY2025_CHECK:
                return f"{namespace}:{tag}", values
    return None


def derive_intangibles():
    """Fallback when no reported tag exists. Uses the balance sheet identity in the latest filings:
    intangibles = total assets - current assets - non-current securities - PP&E - other non-current assets."""
    parts = [("Assets", 1), ("AssetsCurrent", -1), ("MarketableSecuritiesNoncurrent", -1),
             ("PropertyPlantAndEquipmentNet", -1), ("OtherAssetsNoncurrent", -1)]
    values = {}
    for year in (2025, 2026):
        total = 0
        for tag, sign in parts:
            result = get_annual_multi([tag], kind="instant", first_fy=year, last_fy=year)
            if result is None or result.empty:
                total = None
                break
            total += sign * float(result["val"].iloc[0])
        if total is not None:
            values[year] = total
    return values


def add_intangibles(ws):
    """Add the 'Intangible assets, net' row to Raw BS and fill the years that are available."""
    found = find_reported_intangible()
    if found:
        tag_note, values = found
        print(f"Intangible assets: reported tag {tag_note}; fiscal years in SEC data: {sorted(values)}")
    else:
        values = derive_intangibles()
        tag_note = "Derived: total assets - current assets - non-current securities - PP&E - other non-current assets"
        if round(values.get(2025, 0) / 1e6) != INTANGIBLE_FY2025_CHECK:
            print("Intangible assets: no reported tag found and the derived FY2025 figure "
                  f"({ {y: round(v / 1e6) for y, v in values.items()} }) does not match "
                  f"{INTANGIBLE_FY2025_CHECK}. Nothing added.")
            return 0
        print("Intangible assets: no reported tag found, so FY2025-FY2026 are derived from the "
              "balance sheet totals (FY2025 ties to the check figure). FY2021-FY2024 left blank.")

    headers = {str(c.value): c.column for c in ws[1] if c.value is not None}
    year_cols = {y: headers[f"FY{y}"] for y in range(2021, 2027) if f"FY{y}" in headers}
    labels = {str(c.value).strip(): c.row for c in ws["A"] if c.value is not None}

    # Find the row, or add it under the last used row (never inserting or overwriting)
    if INTANGIBLE_LABEL in labels:
        row = labels[INTANGIBLE_LABEL]
    else:
        row = max(labels.values()) + 1
        if any(ws.cell(row, col).value is not None for col in range(1, max(headers.values()) + 2)):
            print(f"Raw BS: row {row} is not empty. Intangible assets not added.")
            return 0
        ws.cell(row, 1, INTANGIBLE_LABEL)

    written = 0
    for year, col in year_cols.items():
        cell = ws.cell(row, col)
        if cell.value is not None:
            continue                                  # never overwrite
        if year in values:
            cell.value = float(values[year]) / 1e6    # USD millions
            written += 1
        else:
            cell.value = MISSING_MARK
        cell._style = copy(ws.cell(row - 1, col)._style)   # same number format as the row above

    notes_col = next((c.column for c in ws[1] if c.value == "SEC tag (sign applied)"), None)
    if notes_col and ws.cell(row, notes_col).value is None:
        ws.cell(row, notes_col, tag_note)

    # FY2025 'Other non-current assets': compare your sheet with the latest filing's recast figure
    if "Other non-current assets" in labels and 2025 in year_cols:
        nca_cell = ws.cell(labels["Other non-current assets"], year_cols[2025])
        latest = get_annual_multi(["OtherAssetsNoncurrent"], kind="instant", first_fy=2025, last_fy=2025)
        if latest is not None and not latest.empty and nca_cell.value is not None:
            latest_value = float(latest["val"].iloc[0]) / 1e6
            if abs(nca_cell.value - latest_value) > 0.5:
                print(f"NOTE: FY2025 Other non-current assets in your sheet = {nca_cell.value:,.0f}; "
                      f"latest filing = {latest_value:,.0f}. The gap should equal FY2025 intangibles.")
                if RECAST_FY2025_OTHER_NCA:
                    nca_cell.value = latest_value
                    print("  Updated the FY2025 cell to the recast figure.")
    return written

# --- Main ------------------------------------------------------------------------
if not os.path.exists(MODEL_FILE):
    raise SystemExit(f"Could not find {MODEL_FILE} in {os.getcwd()}. Nothing written.")

workbook = load_workbook(MODEL_FILE)
absent = [name for name in ITEMS if name not in workbook.sheetnames]
if absent:
    raise SystemExit(f"Sheets not found: {absent}. Nothing written.")

total = 0
for sheet_name, items in ITEMS.items():
    count = fill_sheet(workbook[sheet_name], items)
    print(f"{sheet_name}: {count} {HEADER} value(s) written")
    total += count

total += add_intangibles(workbook["Raw BS"])

if total == 0:
    print(f"Nothing written. If {HEADER} is missing from the SEC data, re-run step1_download.py first.")
else:
    shutil.copy(MODEL_FILE, "Apple_Financial_Model_backup_before_FY2026.xlsx")
    try:
        workbook.save(MODEL_FILE)
        print(f"Saved {MODEL_FILE}.")
    except PermissionError:
        print("Could not save. Close the Excel file and run the script again.")