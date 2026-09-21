"""
Deterministic fake accounts + transactions in Plaid's shape, for `--demo` runs and tests.

Everything here is invented: no real person, bank, employer or merchant is represented,
so this file is safe to publish. Keep it that way — if you want the demo to mirror your
own accounts, do it in a local copy rather than in the committed file.
"""
from __future__ import annotations

import random
from datetime import date, timedelta

DEMO_ITEMS = {
    "item_demo_bank": {"institution": "Demo Bank"},
    "item_demo_cu": {"institution": "Demo Credit Union"},
}

DEMO_ACCOUNTS = [
    {"account_id": "acct_a_checking", "item": "item_demo_bank", "name": "Checking", "mask": "1234", "type": "depository", "subtype": "checking", "owner": "Owner A", "balance": 1840.22},
    {"account_id": "acct_b_spending", "item": "item_demo_bank", "name": "Checking", "mask": "5678", "type": "depository", "subtype": "checking", "owner": "Owner B", "balance": 612.40},
    {"account_id": "acct_j_bills",    "item": "item_demo_bank", "name": "Bills Vault", "mask": "9012", "type": "depository", "subtype": "savings", "owner": "Joint", "balance": 950.00},
    {"account_id": "acct_j_goal",    "item": "item_demo_bank", "name": "Goal Vault", "mask": "3456", "type": "depository", "subtype": "savings", "owner": "Joint", "balance": 18361.94},
    {"account_id": "acct_b_card",     "item": "item_demo_cu", "name": "Cash Rewards Card", "mask": "7890", "type": "credit", "subtype": "credit card", "owner": "Owner B", "balance": 212.55},
]

_MERCHANTS = [
    ("Northside Market", 40, 160, "FOOD_AND_DRINK", "FOOD_AND_DRINK_GROCERIES", "acct_j_bills"),
    ("Valley Superstore", 30, 120, "GENERAL_MERCHANDISE", "GENERAL_MERCHANDISE_SUPERSTORES", "acct_b_card"),
    ("Corner Chicken Co", 9, 24, "FOOD_AND_DRINK", "FOOD_AND_DRINK_FAST_FOOD", "acct_a_checking"),
    ("Rio Burrito", 11, 30, "FOOD_AND_DRINK", "FOOD_AND_DRINK_RESTAURANT", "acct_b_spending"),
    ("Speedway Fuel", 25, 60, "TRANSPORTATION", "TRANSPORTATION_GAS", "acct_a_checking"),
    ("Everything Online", 12, 90, "GENERAL_MERCHANDISE", "GENERAL_MERCHANDISE_ONLINE_MARKETPLACES", "acct_b_card"),
    ("Bullseye Retail", 20, 80, "GENERAL_MERCHANDISE", "GENERAL_MERCHANDISE_DEPARTMENT_STORES", "acct_b_spending"),
    ("Corner Pharmacy", 8, 45, "MEDICAL", "MEDICAL_PHARMACIES_AND_SUPPLEMENTS", "acct_a_checking"),
    ("Grand Cinema", 22, 40, "ENTERTAINMENT", "ENTERTAINMENT_TV_AND_MOVIES", "acct_b_card"),
]

_FIXED = [  # (day, name, amount, primary, detailed, account)
    (1, "Oakwood Property Mgmt Rent", 1105.00, "RENT_AND_UTILITIES", "RENT_AND_UTILITIES_RENT", "acct_j_bills"),
    (11, "City Power & Light", 128.44, "RENT_AND_UTILITIES", "RENT_AND_UTILITIES_GAS_AND_ELECTRICITY", "acct_j_bills"),
    (11, "Demo Wireless", 110.00, "RENT_AND_UTILITIES", "RENT_AND_UTILITIES_TELEPHONE", "acct_j_bills"),
    (11, "Demo Broadband", 52.00, "RENT_AND_UTILITIES", "RENT_AND_UTILITIES_INTERNET_AND_CABLE", "acct_j_bills"),
    (12, "Iron Works Gym", 33.00, "PERSONAL_CARE", "PERSONAL_CARE_GYMS_AND_FITNESS_CENTERS", "acct_j_bills"),
    (15, "Streamtunes", 19.99, "ENTERTAINMENT", "ENTERTAINMENT_MUSIC_AND_AUDIO", "acct_b_card"),
    (15, "Bingeflix", 22.99, "ENTERTAINMENT", "ENTERTAINMENT_TV_AND_MOVIES", "acct_b_card"),
    (16, "Cloud Tools Inc", 20.00, "GENERAL_SERVICES", "GENERAL_SERVICES_OTHER_GENERAL_SERVICES", "acct_a_checking"),
    (20, "Shield Mutual Life", 50.00, "GENERAL_SERVICES", "GENERAL_SERVICES_INSURANCE", "acct_j_bills"),
    (29, "Acme Loan Servicing", 200.00, "LOAN_PAYMENTS", "LOAN_PAYMENTS_STUDENT_LOAN", "acct_b_spending"),
]

_INCOME = [  # (day, name, amount, account)
    (4, "EMPLOYER A PAYROLL", 800.00, "acct_a_checking"),
    (18, "EMPLOYER A PAYROLL", 800.00, "acct_a_checking"),
    (5, "PAYROLL DIRECT DEP", 650.00, "acct_b_spending"),
    (19, "PAYROLL DIRECT DEP", 650.00, "acct_b_spending"),
]

_TRANSFERS = [  # (day, name, amount, primary, detailed, account)
    (6, "Transfer to Bills Vault", 500.00, "TRANSFER_OUT", "TRANSFER_OUT_SAVINGS", "acct_a_checking"),
    (6, "Transfer from Checking", -500.00, "TRANSFER_IN", "TRANSFER_IN_SAVINGS", "acct_j_bills"),
    (6, "Transfer to Bills Vault", 500.00, "TRANSFER_OUT", "TRANSFER_OUT_SAVINGS", "acct_b_spending"),
    (6, "Transfer from Checking", -500.00, "TRANSFER_IN", "TRANSFER_IN_SAVINGS", "acct_j_bills"),
    (20, "Transfer to Goal Vault", 400.00, "TRANSFER_OUT", "TRANSFER_OUT_SAVINGS", "acct_a_checking"),
    (20, "Transfer from Checking", -400.00, "TRANSFER_IN", "TRANSFER_IN_SAVINGS", "acct_j_goal"),
    (25, "CREDIT CARD PAYMENT", 250.00, "LOAN_PAYMENTS", "LOAN_PAYMENTS_CREDIT_CARD_PAYMENT", "acct_b_spending"),
    (25, "Payment - Thank You", -250.00, "LOAN_PAYMENTS", "LOAN_PAYMENTS_CREDIT_CARD_PAYMENT", "acct_b_card"),
]


def demo_accounts_by_item() -> dict[str, list[dict]]:
    out: dict[str, list[dict]] = {}
    for a in DEMO_ACCOUNTS:
        out.setdefault(a["item"], []).append({
            "account_id": a["account_id"], "name": a["name"], "official_name": a["name"], "mask": a["mask"],
            "type": a["type"], "subtype": a["subtype"],
            "balances": {"current": a["balance"], "available": a["balance"]},
        })
    return out


def demo_transactions(start: date, end: date, seed: int = 7) -> list[dict]:
    rng = random.Random(seed)
    txns = []
    n = 0

    def add(d: date, name: str, amount: float, primary: str, detailed: str, acct: str, pending=False):
        nonlocal n
        n += 1
        txns.append({
            "transaction_id": f"demo_{n:05d}", "account_id": acct, "date": d.isoformat(),
            "authorized_date": (d - timedelta(days=1)).isoformat() if rng.random() < 0.5 else d.isoformat(),
            "name": name, "merchant_name": name, "amount": round(amount, 2), "pending": pending,
            "personal_finance_category": {"primary": primary, "detailed": detailed},
        })

    cur = date(start.year, start.month, 1)
    while cur <= end:
        for day, name, amt, pri, det, acct in _FIXED:
            d = cur.replace(day=min(day, 28))
            if start <= d <= end:
                add(d, name, amt, pri, det, acct)
        for day, name, amt, acct in _INCOME:
            d = cur.replace(day=min(day, 28))
            if start <= d <= end:
                add(d, name, -amt, "INCOME", "INCOME_WAGES", acct)
        for day, name, amt, pri, det, acct in _TRANSFERS:
            d = cur.replace(day=min(day, 28))
            if start <= d <= end:
                add(d, name, amt, pri, det, acct)
        for _ in range(rng.randint(14, 22)):
            name, lo, hi, pri, det, acct = rng.choice(_MERCHANTS)
            d = cur.replace(day=rng.randint(1, 28))
            if start <= d <= end:
                add(d, name, rng.uniform(lo, hi), pri, det, acct, pending=(d >= end - timedelta(days=2)))
        cur = (cur.replace(day=28) + timedelta(days=4)).replace(day=1)
    return txns
