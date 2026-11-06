"""
Victory Glass
Forecast Builder V17
Sprint 1

Loads:
    AR Aging Detail export
    Unpaid Bills Detail export

Validates:
    Headers
    Invoice rows
    Bill rows

Outputs:
    AR total
    AP total
"""

from pathlib import Path
import json
import pandas as pd


ROOT = Path(__file__).parent

CONFIG_FILE = ROOT / "config.json"

INPUT_FOLDER = ROOT / "input"

AR_FILE = INPUT_FOLDER / "AR Aging Detail export.xlsx"
AP_FILE = INPUT_FOLDER / "Unpaid Bills Detail export.xlsx"


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


def validate_headers(df, required_headers):

    missing = []

    for header in required_headers:

        if header not in df.columns:
            missing.append(header)

    return missing


def process_ar():

    print("\nLoading AR export...")

    df = pd.read_excel(AR_FILE)

    missing = validate_headers(
        df,
        AR_HEADERS
    )

    if missing:

        raise ValueError(
            f"Missing AR headers: {missing}"
        )

    invoices = df[
        df["Type"] == "Invoice"
    ].copy()

    ar_total = invoices[
        "Open Balance"
    ].sum()

    print(
        f"Invoices Loaded: {len(invoices):,}"
    )

    print(
        f"AR Total: ${ar_total:,.2f}"
    )

    return invoices


def process_ap():

    print("\nLoading AP export...")

    df = pd.read_excel(AP_FILE)

    missing = validate_headers(
        df,
        AP_HEADERS
    )

    if missing:

        raise ValueError(
            f"Missing AP headers: {missing}"
        )

    bills = df[
        df["Type"] == "Bill"
    ].copy()

    ap_total = bills[
        "Open Balance"
    ].sum()

    print(
        f"Bills Loaded: {len(bills):,}"
    )

    print(
        f"AP Total: ${ap_total:,.2f}"
    )

    return bills


def main():

    print("\n======================")
    print("VICTORY GLASS FORECAST")
    print("======================")

    config = load_config()

    print(
        f"\nCurrent Cash: "
        f"${config['current_cash'\]:,.2f}"
    )

    process_ar()

    process_ap()

    print("\nSTATUS: READY")


if __name__ == "__main__":
    main()
