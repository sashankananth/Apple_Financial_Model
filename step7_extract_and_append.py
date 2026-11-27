import os
import shutil
import pandas as pd
from openpyxl import load_workbook
from sec_helpers import get_annual_multi, get_annual_duration, get_annual_instant, us_gaap

YEARS = [2021, 2022, 2023, 2024, 2025]
YEAR_LABELS = [f"FY{y}" for y in YEARS]
MODEL_FILE = "Apple_Financial_Model.xlsx"
TOL = 1_000_000  # checks allow a $1m rounding difference

# --- 1. New items: (label, sheet, kind, candidate tags, sign) ----------------
# sign = -1 flips the SEC value so it matches the printed statement.
NEW_ITEMS = [
    # Balance sheet
    ("Other current assets", "Raw BS", "instant", ["OtherAssetsCurrent"], 1),
    ("Other non-current assets", "Raw BS", "instant", ["OtherAssetsNoncurrent"], 1),
    ("Deferred revenue (current)", "Raw BS", "instant",
     ["ContractWithCustomerLiabilityCurrent", "DeferredRevenueCurrent"], 1),
    ("Deferred revenue (non-current)", "Raw BS", "instant",
     ["ContractWithCustomerLiabilityNoncurrent", "DeferredRevenueNoncurrent"], 1),
    ("Other current liabilities", "Raw BS", "instant", ["OtherLiabilitiesCurrent"], 1),
    ("Other non-current liabilities", "Raw BS", "instant", ["OtherLiabilitiesNoncurrent"], 1),

    # Cash flow - operating
    ("Other non-cash adjustments (incl. deferred income tax)", "Raw CF", "duration", ["OtherNoncashIncomeExpense"], 1),
    ("Chg in accounts receivable", "Raw CF", "duration", ["IncreaseDecreaseInAccountsReceivable"], -1),
    ("Chg in vendor non-trade receivables", "Raw CF", "duration",
     ["IncreaseDecreaseInOtherReceivables"], -1),
    ("Chg in inventories", "Raw CF", "duration", ["IncreaseDecreaseInInventories"], -1),
    ("Chg in other current & non-current assets", "Raw CF", "duration",
     ["IncreaseDecreaseInOtherOperatingAssets"], -1),
    ("Chg in accounts payable", "Raw CF", "duration", ["IncreaseDecreaseInAccountsPayable"], 1),
    ("Chg in other current & non-current liabilities", "Raw CF", "duration",
     ["IncreaseDecreaseInOtherOperatingLiabilities"], 1),

    # Cash flow - investing
    ("Purchases of marketable securities", "Raw CF", "duration",
     ["PaymentsToAcquireAvailableForSaleSecuritiesDebt", "PaymentsToAcquireAvailableForSaleSecurities"], -1),
    ("Proceeds from maturities of marketable securities", "Raw CF", "duration",
     ["ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities",
      "ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecuritiesDebt"], 1),
    ("Proceeds from sales of marketable securities", "Raw CF", "duration",
     ["ProceedsFromSaleOfAvailableForSaleSecuritiesDebt", "ProceedsFromSaleOfAvailableForSaleSecurities"], 1),
    ("Other investing activities", "Raw CF", "duration",
     ["PaymentsForProceedsFromOtherInvestingActivities"], -1),
    # Cash flow - financing
    ("Proceeds from issuance of debt", "Raw CF", "duration", ["ProceedsFromIssuanceOfLongTermDebt"], 1),
    ("Repayments of debt", "Raw CF", "duration", ["RepaymentsOfLongTermDebt"], -1),
    ("Commercial paper, net", "Raw CF", "duration", ["ProceedsFromRepaymentsOfCommercialPaper"], 1),
    ("Taxes paid on net share settlement", "Raw CF", "duration",
     ["PaymentsRelatedToTaxWithholdingForShareBasedCompensation"], -1),
    ("Other financing activities", "Raw CF", "duration",
     ["ProceedsFromPaymentsForOtherFinancingActivities"], 1),
    # Cash reconciliation
    ("Net change in cash", "Raw CF", "duration",
     ["CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseExcludingExchangeRateEffect",
      "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect",
      "CashAndCashEquivalentsPeriodIncreaseDecrease"], 1),
    ("FX effect on cash", "Raw CF", "duration",
     ["EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
      "EffectOfExchangeRateOnCashAndCashEquivalents"], 1),
]

# --- 2. Existing items, needed ONLY for the checks (not exported again) -------
EXISTING = {
    "Cash": ("CashAndCashEquivalentsAtCarryingValue", "instant"),
    "Current securities": ("MarketableSecuritiesCurrent", "instant"),
    "Accounts receivable": ("AccountsReceivableNetCurrent", "instant"),
    "Vendor receivables": ("NontradeReceivablesCurrent", "instant"),
    "Inventories": ("InventoryNet", "instant"),
    "Total current assets": ("AssetsCurrent", "instant"),
    "Non-current securities": ("MarketableSecuritiesNoncurrent", "instant"),
    "PP&E": ("PropertyPlantAndEquipmentNet", "instant"),
    "Total assets": ("Assets", "instant"),
    "Accounts payable": ("AccountsPayableCurrent", "instant"),
    "Commercial paper": ("CommercialPaper", "instant"),
    "Current term debt": ("LongTermDebtCurrent", "instant"),
    "Total current liabilities": ("LiabilitiesCurrent", "instant"),
    "Non-current term debt": ("LongTermDebtNoncurrent", "instant"),
    "Total liabilities": ("Liabilities", "instant"),
    "Net income": ("NetIncomeLoss", "duration"),
    "D&A": ("DepreciationDepletionAndAmortization", "duration"),
    "Share-based comp": ("ShareBasedCompensation", "duration"),
    "Operating cash flow": ("NetCashProvidedByUsedInOperatingActivities", "duration"),
    "Capex": ("PaymentsToAcquirePropertyPlantAndEquipment", "duration"),
    "Investing cash flow": ("NetCashProvidedByUsedInInvestingActivities", "duration"),
    "Dividends paid": ("PaymentsOfDividends", "duration"),
    "Share repurchases": ("PaymentsForRepurchaseOfCommonStock", "duration"),
    "Financing cash flow": ("NetCashProvidedByUsedInFinancingActivities", "duration"),
}


def to_series(result):
    """Turn a helper result into one value per fiscal year (blank if missing)."""
    return result.set_index("fiscal_year")["val"].reindex(YEARS)


# --- 3. Pull the data ---------------------------------------------------------
new, notes, sheet_of = {}, {}, {}
for label, sheet, kind, tags, sign in NEW_ITEMS:
    result = get_annual_multi(tags, kind)
    if result is None:
        print(f"NOT AVAILABLE (not exported): {label}")
        continue
    new[label] = to_series(result) * sign
    notes[label] = " / ".join(result["tag"].unique()) + ("  (sign flipped)" if sign == -1 else "")
    sheet_of[label] = sheet
    missing = [y for y in YEARS if y not in result["fiscal_year"].tolist()]
    if missing:
        print(f"NOTE: '{label}' has no value for {missing}")

# Document why non-current deferred revenue is unavailable
points = us_gaap.get("DeferredRevenueNoncurrent", {}).get("units", {}).get("USD", [])
print("DeferredRevenueNoncurrent periods in the file (end, form):",
      sorted({(p["end"], p["form"]) for p in points}))

old = {}
for name, (tag, kind) in EXISTING.items():
    func = get_annual_duration if kind == "duration" else get_annual_instant
    result = func(tag)
    if result is None:
        raise SystemExit(f"Could not load existing item {name} ({tag})")
    old[name] = to_series(result)
# --- Beginning and ending cash (cash flow statement basis, incl. restricted cash) ---
CASH_TAG = "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"
BEGIN_LABEL = "Cash, cash equivalents & restricted cash - beginning of year"
END_LABEL = "Cash, cash equivalents & restricted cash - end of year"

# FY2020 is needed because it is the opening balance of FY2021
cash = get_annual_instant(CASH_TAG, first_fy=2020, last_fy=2025)
if cash is None:
    print(f"NOT AVAILABLE (not exported): {END_LABEL}")
else:
    ending = cash.set_index("fiscal_year")["val"].reindex(range(2020, 2026))
    if ending.isna().any():
        print("NOTE: cash balance missing for fiscal years", ending[ending.isna()].index.tolist())
    new[BEGIN_LABEL] = ending.shift(1).reindex(YEARS)   # last year's ending balance
    new[END_LABEL] = ending.reindex(YEARS)
    for label in (BEGIN_LABEL, END_LABEL):
        sheet_of[label] = "Raw CF"
    notes[BEGIN_LABEL] = CASH_TAG + " (prior year-end)"
    notes[END_LABEL] = CASH_TAG

# --- 4. Validation ------------------------------------------------------------
def check(label, left, right):
    """Pass only if left and right match in every year (a blank counts as a failure)."""
    gap = left - right
    bad = gap.isna() | (gap.abs() > TOL)
    if bad.any():
        print(f"FAILED {label}: years {gap[bad].index.tolist()}, gap (USD m) {(gap[bad] / 1e6).round(1).tolist()}")
        return False
    print(f"OK     {label}")
    return True


print("\n--- Balance sheet checks ---")
bs_ok = True
try:
    bs_ok &= check("Current assets add up",
                   old["Cash"] + old["Current securities"] + old["Accounts receivable"]
                   + old["Vendor receivables"] + old["Inventories"] + new["Other current assets"],
                   old["Total current assets"])
    bs_ok &= check("Total assets add up",
                   old["Total current assets"] + old["Non-current securities"] + old["PP&E"]
                   + new["Other non-current assets"], old["Total assets"])
    bs_ok &= check("Current liabilities add up",
                   old["Accounts payable"] + new["Other current liabilities"]
                   + new["Deferred revenue (current)"] + old["Commercial paper"]
                   + old["Current term debt"], old["Total current liabilities"])
    bs_ok &= check("Total liabilities add up",
                   old["Total current liabilities"] + old["Non-current term debt"]
                   + new["Other non-current liabilities"], old["Total liabilities"])
except KeyError as error:
    print(f"Balance sheet checks skipped, missing item: {error}")
    bs_ok = False

print("\n--- Cash flow checks ---")
cf_ok = True
try:
    # Operating: everything except "Other non-cash" is known, so work out what it must be
    known = (old["Net income"] + old["D&A"] + old["Share-based comp"]
             + new["Chg in accounts receivable"] + new["Chg in vendor non-trade receivables"]
             + new["Chg in inventories"] + new["Chg in other current & non-current assets"]
             + new["Chg in accounts payable"] + new["Chg in other current & non-current liabilities"])
    needed = old["Operating cash flow"] - known
    raw = new["Other non-cash adjustments (incl. deferred income tax)"]
    if ((needed - raw).abs() <= TOL).all():
        print("OK     Operating cash flow reconciles (Other non-cash sign kept)")
    elif ((needed + raw).abs() <= TOL).all():
        new["Other non-cash adjustments (incl. deferred income tax)"] = -raw
        notes["Other non-cash adjustments (incl. deferred income tax)"] += "  (sign flipped)"
        print("OK     Operating cash flow reconciles (Other non-cash sign flipped)")
    else:
        print("FAILED Operating cash flow does not reconcile")
        print("   'Other non-cash' needed (USD m):", (needed / 1e6).round(1).tolist())
        print("   SEC value (USD m):              ", (raw / 1e6).round(1).tolist())
        cf_ok = False

    cf_ok &= check("Investing components add up",
                   new["Purchases of marketable securities"]
                   + new["Proceeds from maturities of marketable securities"]
                   + new["Proceeds from sales of marketable securities"]
                   - old["Capex"] + new["Other investing activities"],
                   old["Investing cash flow"])
    cf_ok &= check("Financing components add up",
                   new["Proceeds from issuance of debt"] + new["Repayments of debt"]
                   + new["Commercial paper, net"] + new["Taxes paid on net share settlement"]
                   + new["Other financing activities"]
                   - old["Dividends paid"] - old["Share repurchases"],
                   old["Financing cash flow"])
    cf_ok &= check("Operating + investing + financing = net change in cash",
                   old["Operating cash flow"] + old["Investing cash flow"] + old["Financing cash flow"],
                   new["Net change in cash"])
    cf_ok &= check("Beginning cash + net change = ending cash",
                   new[BEGIN_LABEL] + new["Net change in cash"], new[END_LABEL])
    print("Info: restricted cash (CF ending cash less BS cash), USD m:",
          ((new[END_LABEL] - old["Cash"]) / 1e6).round(0).tolist())
    
except KeyError as error:
    print(f"Cash flow checks skipped, missing item: {error}")
    cf_ok = False

print("\n--- Tags used ---")
for label in new:
    print(f"{label:50} {notes[label]}")


# --- 5. Append to the existing Raw sheets --------------------------------------
def append_rows(ws, rows):
    """Add rows under the existing data. Never overwrites, never duplicates.
    rows = list of (label, series in USD, note). Returns how many rows were added."""
    # Find the column of each fiscal year from the header row
    year_col = {cell.value: cell.column for cell in ws[1] if cell.value in YEAR_LABELS}
    if len(year_col) != len(YEAR_LABELS):
        print(f"  {ws.title}: FY2021-FY2025 headers not found in row 1. Nothing written.")
        return 0
    note_col = max(year_col.values()) + 1

    labels_in_sheet = {str(c.value).strip() for c in ws["A"] if c.value is not None}
    last_row = max(c.row for c in ws["A"] if c.value is not None)
    added = 0
    for label, series, note in rows:
        if label in labels_in_sheet:
            print(f"  {ws.title}: '{label}' already exists - skipped")
            continue
        row = last_row + 1
        if any(ws.cell(row, col).value is not None for col in range(1, note_col + 1)):
            print(f"  {ws.title}: row {row} is not empty. Stopped to avoid overwriting.")
            break
        ws.cell(row, 1, label)
        for year in YEARS:
            if pd.notna(series[year]):
                ws.cell(row, year_col[f"FY{year}"], float(series[year]) / 1e6)  # USD millions
        ws.cell(row, note_col, note)
        last_row = row
        added += 1
    if added and ws.cell(1, note_col).value is None:
        ws.cell(1, note_col, "SEC tag (sign applied)")
    return added


print("\n--- Excel ---")
if not os.path.exists(MODEL_FILE):
    raise SystemExit(f"Could not find {MODEL_FILE} in {os.getcwd()}. Nothing written.")

workbook = load_workbook(MODEL_FILE)
total_added = 0
for sheet_name, ok in (("Raw BS", bs_ok), ("Raw CF", cf_ok)):
    if sheet_name not in workbook.sheetnames:
        print(f"Sheet '{sheet_name}' not found - skipped.")
    elif not ok:
        print(f"{sheet_name}: checks failed, nothing written.")
    else:
        rows = [(label, new[label], notes[label]) for label in new if sheet_of[label] == sheet_name]
        added = append_rows(workbook[sheet_name], rows)
        print(f"{sheet_name}: {added} row(s) added")
        total_added += added

if total_added == 0:
    print("Nothing new to write. The workbook was not touched.")
else:
    shutil.copy(MODEL_FILE, "Apple_Financial_Model_backup.xlsx")  # copy of the file before this save
    try:
        workbook.save(MODEL_FILE)
        print(f"Saved {MODEL_FILE} ({total_added} new rows).")
    except PermissionError:
        print("Could not save. Close the Excel file and run the script again.")