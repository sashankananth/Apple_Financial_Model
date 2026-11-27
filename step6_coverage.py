from sec_helpers import get_annual_multi, data

# (name, kind, [candidate tags in priority order])
NEW_ITEMS = [
    # Balance sheet
    ("Other current assets", "instant", ["OtherAssetsCurrent"]),
    ("Other non-current assets", "instant", ["OtherAssetsNoncurrent"]),
    ("Deferred revenue (current)", "instant",
     ["ContractWithCustomerLiabilityCurrent", "DeferredRevenueCurrent"]),
    ("Deferred revenue (non-current)", "instant",
     ["ContractWithCustomerLiabilityNoncurrent", "DeferredRevenueNoncurrent"]),
    ("Other current liabilities", "instant", ["OtherLiabilitiesCurrent"]),
    ("Other non-current liabilities", "instant", ["OtherLiabilitiesNoncurrent"]),
    # Cash flow - operating
    ("Chg in accounts receivable", "duration", ["IncreaseDecreaseInAccountsReceivable"]),
    ("Chg in vendor non-trade receivables", "duration", ["IncreaseDecreaseInOtherReceivables"]),
    ("Chg in inventories", "duration", ["IncreaseDecreaseInInventories"]),
    ("Chg in accounts payable", "duration", ["IncreaseDecreaseInAccountsPayable"]),
    ("Chg in other operating assets", "duration", ["IncreaseDecreaseInOtherOperatingAssets"]),
    ("Chg in other operating liabilities", "duration", ["IncreaseDecreaseInOtherOperatingLiabilities"]),
    ("Deferred income tax", "duration", ["DeferredIncomeTaxExpenseBenefit"]),
    ("Other non-cash adjustments", "duration", ["OtherNoncashIncomeExpense"]),
    # Cash flow - investing
    ("Purchases of marketable securities", "duration",
     ["PaymentsToAcquireAvailableForSaleSecuritiesDebt", "PaymentsToAcquireAvailableForSaleSecurities"]),
    ("Maturities of marketable securities", "duration",
     ["ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities",
      "ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecuritiesDebt"]),
    ("Sales of marketable securities", "duration",
     ["ProceedsFromSaleOfAvailableForSaleSecuritiesDebt", "ProceedsFromSaleOfAvailableForSaleSecurities"]),
    ("Acquisitions", "duration", ["PaymentsToAcquireBusinessesNetOfCashAcquired"]),
    ("Other investing activities", "duration", ["PaymentsForProceedsFromOtherInvestingActivities"]),
    # Cash flow - financing
    ("Proceeds from debt issuance", "duration", ["ProceedsFromIssuanceOfLongTermDebt"]),
    ("Repayments of debt", "duration", ["RepaymentsOfLongTermDebt"]),
    ("Commercial paper, net", "duration", ["ProceedsFromRepaymentsOfCommercialPaper"]),
    ("Taxes on net share settlement", "duration",
     ["PaymentsRelatedToTaxWithholdingForShareBasedCompensation"]),
    # Cash reconciliation
    ("Net change in cash", "duration",
     ["CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseExcludingExchangeRateEffect",
      "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect",
      "CashAndCashEquivalentsPeriodIncreaseDecrease"]),
    ("FX effect on cash", "duration",
     ["EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
      "EffectOfExchangeRateOnCashAndCashEquivalents"]),
]

print(f"{'Line item':38} {'Found':>5}  Missing   Tag(s) used")
print("-" * 110)
for name, kind, candidates in NEW_ITEMS:
    result = get_annual_multi(candidates, kind)
    found = [] if result is None else result["fiscal_year"].tolist()
    missing = [y for y in range(2021, 2026) if y not in found]
    tags_used = "-" if result is None else ", ".join(sorted(result["tag"].unique()))
    flag = "  <-- CHECK" if missing else ""
    print(f"{name:38} {len(found):>3}/5  {str(missing):9} {tags_used}{flag}")

# --- Search ALL namespaces (us-gaap, aapl, dei ...) for tags containing a keyword ---
KEYWORDS = ["Receivable", "NonTrade", "Nontrade", "DeferredRevenue", "ContractWithCustomerLiability"]
print("\nTags containing the keywords (namespace: tag):")
for namespace, tags in data["facts"].items():
    for tag in tags:
        if any(word.lower() in tag.lower() for word in KEYWORDS):
            print(f"  {namespace}: {tag}")