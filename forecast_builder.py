# Victory Glass Forecast Builder
"""
Victory Glass
Forecast Builder V17
Sprint 1

Purpose:
- Validate QuickBooks exports
- Verify AR and AP imports
- Confirm totals
"""

from pathlib import Path
import json
import pandas as pd


ROOT = Path(__file__).parent

INPUT_FOLDER = ROOT / "input"

AR_FILE = INPUT_FOLDER / "AR Aging Detail export.xlsx"
AP_FILE = INPUT_FOLDER / "Unpaid Bills Detail export.xlsx"

CONFIG_FILE = ROOT / "config.json"


AR_HEADERS = [
    "Type",
    "Date",
    "Num",
    "P. O. #",
    "Name",
    "Terms",
    "Due Date",
    "Aging",
    "Open Balance"
]

AP_HEADERS = [
    "Type",
    "Date",
    "Num",
    "Name",
    "Terms",
    "Due Date",
    "Open Balance"
]


def load_config():

    with open(CONFIG_FILE, "r") as f:
        return json.load(f)


def validate_headers(df, required):

    missing = []

    for col in required:
        if col not in df.columns:
            missing.append(col)

    return missing


def process_ar():

    print("\nLoading AR export...")

    df = pd.read_excel(AR_FILE)

    missing = validate_headers(df, AR_HEADERS)

    if missing:
        raise ValueError(
            f"AR headers missing: {missing}"
        )

    invoices = df[
        df["Type"] == "Invoice"
    ].copy()

    ar_total = invoices[
        "Open Balance"
    ].sum()

    return invoices, ar_total


def process_ap():

    print("\nLoading AP export...")

    df = pd.read_excel(AP_FILE)

    missing = validate_headers(df, AP_HEADERS)

    if missing:
        raise ValueError(
            f"AP headers missing: {missing}"
        )

    bills = df[
        df["Type"] == "Bill"
    ].copy()

    ap_total = bills[
        "Open Balance"
    ].sum()

    return bills, ap_total


def main():

    config = load_config()

    print("\n--------------------------------")
    print("VICTORY GLASS CASH FORECAST")
    print("--------------------------------")

    print(
        f"Current Cash: "
        f"${config['current_cash']:,.2f}"
    )

    ar_df, ar_total = process_ar()

    ap_df, ap_total = process_ap()

    print("\nAR RESULTS")

    print(
        f"Invoices Loaded: "
        f"{len(ar_df):,}"
    )

    print(
        f"AR Total: "
        f"${ar_total:,.2f}"
    )

    print("\nAP RESULTS")

    print(
        f"Bills Loaded: "
        f"{len(ap_df):,}"
    )

    print(
        f"AP Total: "
        f"${ap_total:,.2f}"
    )

    print("\nSTATUS: READY")


if __name__ == "__main__":
    main()
