import pandas as pd
from sec_helpers import get_annual_duration, get_annual_instant

YEARS = [2021, 2022, 2023, 2024, 2025]

# (name, tag, kind, unit)
INCOME = [
    ("Net sales", "RevenueFromContractWithCustomerExcludingAssessedTax", "duration", "USD"),
    ("Cost of sales", "CostOfGoodsAndServicesSold", "duration", "USD"),
    ("Gross profit", "GrossProfit", "duration", "USD"),
    ("R&D", "ResearchAndDevelopmentExpense", "duration", "USD"),
    ("SG&A", "SellingGeneralAndAdministrativeExpense", "duration", "USD"),
    ("Total operating expenses", "OperatingExpenses", "duration", "USD"),
    ("Operating income", "OperatingIncomeLoss", "duration", "USD"),
    ("Other income/(expense)", "NonoperatingIncomeExpense", "duration", "USD"),
    ("Income before taxes",
     "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
     "duration", "USD"),
    ("Income tax expense", "IncomeTaxExpenseBenefit", "duration", "USD"),
    ("Net income", "NetIncomeLoss", "duration", "USD"),
    ("Diluted EPS", "EarningsPerShareDiluted", "duration", "USD/shares"),
    ("Diluted shares", "WeightedAverageNumberOfDilutedSharesOutstanding", "duration", "shares"),
]

BALANCE = [
    ("Cash and equivalents", "CashAndCashEquivalentsAtCarryingValue", "instant", "USD"),
    ("Current marketable securities", "MarketableSecuritiesCurrent", "instant", "USD"),
    ("Accounts receivable", "AccountsReceivableNetCurrent", "instant", "USD"),
    ("Vendor non-trade receivables", "NontradeReceivablesCurrent", "instant", "USD"),
    ("Inventories", "InventoryNet", "instant", "USD"),
    ("Total current assets", "AssetsCurrent", "instant", "USD"),
    ("Non-current marketable securities", "MarketableSecuritiesNoncurrent", "instant", "USD"),
    ("PP&E, net", "PropertyPlantAndEquipmentNet", "instant", "USD"),
    ("Total assets", "Assets", "instant", "USD"),
    ("Accounts payable", "AccountsPayableCurrent", "instant", "USD"),
    ("Commercial paper", "CommercialPaper", "instant", "USD"),
    ("Current term debt", "LongTermDebtCurrent", "instant", "USD"),
    ("Total current liabilities", "LiabilitiesCurrent", "instant", "USD"),
    ("Non-current term debt", "LongTermDebtNoncurrent", "instant", "USD"),
    ("Total liabilities", "Liabilities", "instant", "USD"),
    ("Shareholders' equity", "StockholdersEquity", "instant", "USD"),
]

CASHFLOW = [
    ("Net income", "NetIncomeLoss", "duration", "USD"),
    ("D&A", "DepreciationDepletionAndAmortization", "duration", "USD"),
    ("Share-based compensation", "ShareBasedCompensation", "duration", "USD"),
    ("Operating cash flow", "NetCashProvidedByUsedInOperatingActivities", "duration", "USD"),
    ("Capex", "PaymentsToAcquirePropertyPlantAndEquipment", "duration", "USD"),
    ("Investing cash flow", "NetCashProvidedByUsedInInvestingActivities", "duration", "USD"),
    ("Dividends paid", "PaymentsOfDividends", "duration", "USD"),
    ("Share repurchases", "PaymentsForRepurchaseOfCommonStock", "duration", "USD"),
    ("Financing cash flow", "NetCashProvidedByUsedInFinancingActivities", "duration", "USD"),
]


def build_table(items):
    """Make a table: one row per line item, one column per fiscal year."""
    rows = {}
    for name, tag, kind, unit in items:
        func = get_annual_duration if kind == "duration" else get_annual_instant
        result = func(tag, unit=unit)
        if result is None:
            print(f"WARNING: no data for {name}")
            continue
        # turn the result into {2021: value, 2022: value, ...}
        rows[name] = result.set_index("fiscal_year")["val"]
    table = pd.DataFrame(rows).T          # .T flips so items are rows
    return table.reindex(columns=YEARS)   # force columns into FY2021-FY2025 order


income = build_table(INCOME)
balance = build_table(BALANCE)
cashflow = build_table(CASHFLOW)


# --- Validation: identities that must hold ------------------------------------
def check(label, left, right, tolerance=1_000_000):
    """Compare two rows year by year. Allow a $1m tolerance for rounding."""
    diff = (left - right).abs()
    bad = diff[diff > tolerance]
    if bad.empty:
        print(f"OK     {label}")
    else:
        print(f"FAILED {label} in years: {bad.index.tolist()}")


check("Net sales - Cost of sales = Gross profit",
      income.loc["Net sales"] - income.loc["Cost of sales"], income.loc["Gross profit"])
check("Gross profit - Operating expenses = Operating income",
      income.loc["Gross profit"] - income.loc["Total operating expenses"],
      income.loc["Operating income"])
check("Operating income + Other income = Income before taxes",
      income.loc["Operating income"] + income.loc["Other income/(expense)"],
      income.loc["Income before taxes"])
check("Income before taxes - Tax = Net income",
      income.loc["Income before taxes"] - income.loc["Income tax expense"],
      income.loc["Net income"])
check("Total assets = Total liabilities + Equity",
      balance.loc["Total assets"],
      balance.loc["Total liabilities"] + balance.loc["Shareholders' equity"])
check("Income-statement net income = Cash-flow net income",
      income.loc["Net income"], cashflow.loc["Net income"])

# --- Show the tables ------------------------------------------------------------
pd.options.display.float_format = "{:,.2f}".format
print("\nINCOME STATEMENT\n", income / 1e6)
print("\nBALANCE SHEET\n", balance / 1e6)
print("\nCASH FLOW\n", cashflow / 1e6)

# --- Export to Excel: one worksheet per statement ---------------------------------
import os
import shutil
from openpyxl import load_workbook

# --- Convert to millions (EPS stays in USD, shares in millions) ---------------
def to_millions(table, keep_as_is=()):
    out = table.copy()
    for row in out.index:
        if row not in keep_as_is:
            out.loc[row] = out.loc[row] / 1e6
    return out

income_out = to_millions(income, keep_as_is=("Diluted EPS",))  # shares become millions
balance_out = to_millions(balance)
cashflow_out = to_millions(cashflow)

for table in (income_out, balance_out, cashflow_out):
    table.columns = [f"FY{year}" for year in YEARS]
    table.index.name = "USD millions (EPS in USD, shares in millions)"

# --- Export into the existing model workbook ----------------------------------
MODEL_FILE = "Apple_Financial_Model.xlsx"   # must be in the same folder as this script
NEW_SHEETS = {
    "Raw IS": income_out,
    "Raw BS": balance_out,
    "Raw CF": cashflow_out,
}

if not os.path.exists(MODEL_FILE):
    print(f"Could not find {MODEL_FILE} in {os.getcwd()}. Nothing was written.")
else:
    # Safety check: stop if any of the new sheet names already exist
    workbook = load_workbook(MODEL_FILE, read_only=True)
    existing_sheets = workbook.sheetnames
    workbook.close()
    clashes = [name for name in NEW_SHEETS if name in existing_sheets]

    if clashes:
        print(f"These sheets already exist, so nothing was written: {clashes}")
    else:
        # Safety copy, in case anything goes wrong
        shutil.copy(MODEL_FILE, "Apple_Financial_Model_backup.xlsx")
        try:
            with pd.ExcelWriter(MODEL_FILE, engine="openpyxl", mode="a") as writer:
                for sheet_name, table in NEW_SHEETS.items():
                    table.to_excel(writer, sheet_name=sheet_name)
            print(f"Added {list(NEW_SHEETS)} to {MODEL_FILE}")
        except PermissionError:
            print("Could not save. Close the Excel file and run the script again.")
        except Exception as error:
            print(f"Excel export failed: {error}")