"""Assign one of our spreadsheet categories to a transaction."""
from __future__ import annotations

import re


class Categorizer:
    def __init__(self, cfg: dict):
        self.categories = list(cfg["categories"].keys())
        self.rules = [(re.compile(r["pattern"], re.IGNORECASE), r["category"]) for r in cfg.get("merchant_rules", [])]
        pmap = cfg.get("plaid_category_map", {})
        self.detailed = {k: v for k, v in pmap.get("detailed", {}).items() if not k.startswith("_")}
        self.primary = {k: v for k, v in pmap.get("primary", {}).items() if not k.startswith("_")}

    def categorize(self, txn: dict) -> str:
        """txn is our normalized ledger record (see ledger.normalize_plaid_txn)."""
        text = " ".join(filter(None, [txn.get("name"), txn.get("merchant"), txn.get("original_description")]))
        for rx, cat in self.rules:
            if rx.search(text):
                return self._valid(cat)

        det = txn.get("pfc_detailed") or ""
        pri = txn.get("pfc_primary") or ""
        if det in self.detailed:
            return self._valid(self.detailed[det])
        if pri in self.primary:
            cat = self.primary[pri]
            # money coming back on a credit card that Plaid didn't call a transfer/income = refund
            if txn["amount"] < 0 and txn.get("account_type") == "credit" and cat not in ("Transfer", "Income"):
                return "Refund"
            return self._valid(cat)

        if txn["amount"] < 0:
            return "Refund" if txn.get("account_type") == "credit" else "Income"
        return "Other"

    def _valid(self, cat: str) -> str:
        return cat if cat in self.categories else "Other"
