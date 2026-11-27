import json
import pandas as pd

# Load the saved SEC file once, when this helper file is imported
with open("apple_companyfacts.json") as file:
    data = json.load(file)

us_gaap = data["facts"]["us-gaap"]


def _get_points(tag, unit="USD"):
    """Return the data points for a tag as a table, or None if missing."""
    if tag not in us_gaap:
        return None
    points = us_gaap[tag]["units"].get(unit)
    if not points:
        return None
    df = pd.DataFrame(points)
    return df[df["form"] == "10-K"].copy()  # annual reports only


def get_annual_duration(tag, first_fy=2021, last_fy=2025, unit="USD"):
    """Annual values for income statement / cash flow tags (have a start date)."""
    df = _get_points(tag, unit)
    if df is None or df.empty:
        return None

    df["start"] = pd.to_datetime(df["start"])
    df["end"] = pd.to_datetime(df["end"])
    df["days"] = (df["end"] - df["start"]).dt.days
    df = df[df["days"].between(350, 380)]          # full-year periods only

    df["filed"] = pd.to_datetime(df["filed"])
    df = df.sort_values("filed").drop_duplicates(subset="end", keep="last")

    df["fiscal_year"] = df["end"].dt.year
    df = df[df["fiscal_year"].between(first_fy, last_fy)]
    return df[["fiscal_year", "end", "val", "filed"]].sort_values("fiscal_year")


def get_annual_instant(tag, first_fy=2021, last_fy=2025, unit="USD"):
    """Year-end values for balance sheet tags (only an end date, no start)."""
    df = _get_points(tag, unit)
    if df is None or df.empty:
        return None

    df["end"] = pd.to_datetime(df["end"])
    df["filed"] = pd.to_datetime(df["filed"])
    df = df.sort_values("filed").drop_duplicates(subset="end", keep="last")

    df["fiscal_year"] = df["end"].dt.year
    df = df[df["fiscal_year"].between(first_fy, last_fy)]
    return df[["fiscal_year", "end", "val", "filed"]].sort_values("fiscal_year")

def get_annual_multi(candidates, kind="duration", unit="USD", first_fy=2021, last_fy=2025):
    """Try several candidate tags. For each fiscal year, use the first candidate
    (in the order given) that has a value. Also reports which tag was used."""
    func = get_annual_duration if kind == "duration" else get_annual_instant
    pieces = []
    for priority, tag in enumerate(candidates):
        result = func(tag, first_fy, last_fy, unit)
        if result is not None and not result.empty:
            result = result.copy()
            result["tag"] = tag
            result["priority"] = priority   # lower number = preferred tag
            pieces.append(result)
    if not pieces:
        return None
    combined = pd.concat(pieces)
    combined = combined.sort_values("priority").drop_duplicates(subset="fiscal_year", keep="first")
    return combined.sort_values("fiscal_year")[["fiscal_year", "val", "tag"]]