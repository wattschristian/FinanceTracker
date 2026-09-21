"""
Local persistence. Three files live next to the script:

  state.json   - linked Plaid items (cursors, institution names), account metadata + latest balances
  ledger.json  - every transaction we have ever received, keyed by Plaid transaction_id
  config.json  - user-editable settings (owners, categories, rules, account owners)

Plaid access tokens are stored in the OS credential store via `keyring` when it is
available (Windows Credential Manager / macOS Keychain / Secret Service); otherwise
they fall back to state.json. Either way: keep this folder private.
"""
from __future__ import annotations

import json
import os
from datetime import date, datetime
from pathlib import Path

try:
    import keyring  # optional
    _KEYRING = True
except Exception:  # pragma: no cover
    keyring = None
    _KEYRING = False

KEYRING_SERVICE = "household-finance-sync"


def _load_json(path: Path, default):
    if not path.exists():
        return default
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _save_json(path: Path, data) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with tmp.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, sort_keys=True)
    os.replace(tmp, path)


class Store:
    def __init__(self, folder: Path):
        self.folder = folder
        self.config_path = folder / "config.json"
        self.state_path = folder / "state.json"
        self.ledger_path = folder / "ledger.json"
        self.config = _load_json(self.config_path, None)
        if self.config is None:
            raise FileNotFoundError(f"config.json not found in {folder}")
        self.state = _load_json(self.state_path, {"items": {}, "accounts": {}, "last_sync": None})
        self.ledger = _load_json(self.ledger_path, {"transactions": {}})

    # ------------------------------------------------------------- persistence
    def save(self) -> None:
        _save_json(self.config_path, self.config)
        _save_json(self.state_path, self.state)
        _save_json(self.ledger_path, self.ledger)

    # ------------------------------------------------------------------ tokens
    def set_access_token(self, item_id: str, token: str) -> None:
        if _KEYRING:
            try:
                keyring.set_password(KEYRING_SERVICE, item_id, token)
                self.state["items"][item_id]["token_storage"] = "keyring"
                self.state["items"][item_id].pop("access_token", None)
                return
            except Exception:
                pass
        self.state["items"][item_id]["token_storage"] = "file"
        self.state["items"][item_id]["access_token"] = token

    def get_access_token(self, item_id: str) -> str | None:
        item = self.state["items"].get(item_id, {})
        if item.get("token_storage") == "keyring" and _KEYRING:
            try:
                tok = keyring.get_password(KEYRING_SERVICE, item_id)
                if tok:
                    return tok
            except Exception:
                pass
        return item.get("access_token")

    # ---------------------------------------------------------------- accounts
    def upsert_accounts(self, item_id: str, plaid_accounts: list[dict], as_of: datetime) -> list[str]:
        """Record account metadata + balances. Returns account_ids never seen before."""
        new_ids = []
        for a in plaid_accounts:
            aid = a["account_id"]
            if aid not in self.state["accounts"]:
                new_ids.append(aid)
            bal = a.get("balances") or {}
            self.state["accounts"][aid] = {
                "item_id": item_id,
                "name": a.get("name") or a.get("official_name") or "Account",
                "official_name": a.get("official_name"),
                "mask": a.get("mask"),
                "type": a.get("type"),
                "subtype": a.get("subtype"),
                "balance_current": bal.get("current"),
                "balance_available": bal.get("available"),
                "balance_as_of": as_of.strftime("%Y-%m-%d %H:%M"),
            }
        return new_ids

    def account_owner(self, account_id: str) -> str:
        return self.config["accounts"].get(account_id, {}).get("owner", "Joint")

    def account_label(self, account_id: str) -> str:
        acct = self.state["accounts"].get(account_id, {})
        custom = self.config["accounts"].get(account_id, {}).get("label")
        if custom:
            return custom
        inst = self.state["items"].get(acct.get("item_id", ""), {}).get("institution", "")
        mask = f" ••{acct['mask']}" if acct.get("mask") else ""
        return f"{inst} {acct.get('name', 'Account')}{mask}".strip()

    def account_hidden(self, account_id: str) -> bool:
        return bool(self.config["accounts"].get(account_id, {}).get("hidden", False))

    # ------------------------------------------------------------ transactions
    @staticmethod
    def normalize_plaid_txn(t: dict, account_type: str | None) -> dict:
        pfc = t.get("personal_finance_category") or {}
        charge_date = t.get("authorized_date") or t.get("date")
        return {
            "id": t["transaction_id"],
            "account_id": t["account_id"],
            "account_type": account_type,
            "date": charge_date,                    # charge/authorized date (what we sort by)
            "posted_date": t.get("date"),
            "name": t.get("name") or "",
            "merchant": t.get("merchant_name"),
            "original_description": t.get("original_description"),
            "amount": float(t.get("amount") or 0.0),  # Plaid: positive = money out
            "pending": bool(t.get("pending")),
            "pfc_primary": pfc.get("primary"),
            "pfc_detailed": pfc.get("detailed"),
            "removed": False,
        }

    def apply_sync(self, added: list[dict], modified: list[dict], removed: list[dict], categorizer) -> dict:
        txns = self.ledger["transactions"]
        counts = {"added": 0, "modified": 0, "removed": 0, "pending_skipped": 0}
        start = self.config.get("history_start_date")
        for t in added + modified:
            acct_type = self.state["accounts"].get(t["account_id"], {}).get("type")
            rec = self.normalize_plaid_txn(t, acct_type)
            if start and rec["date"] and rec["date"] < start:
                continue
            existing = txns.get(rec["id"])
            if existing:
                rec["category_override"] = existing.get("category_override")
                counts["modified"] += 1
            else:
                counts["added"] += 1
            rec["category"] = categorizer.categorize(rec)
            txns[rec["id"]] = rec
        for r in removed:
            tid = r.get("transaction_id")
            if tid in txns:
                txns[tid]["removed"] = True
                counts["removed"] += 1
        counts["pending_skipped"] = sum(1 for t in txns.values() if t.get("pending") and not t.get("removed"))
        return counts

    def apply_category_overrides(self, overrides: dict[str, str]) -> int:
        """overrides: transaction_id -> category the user typed into the sheet."""
        n = 0
        valid = set(self.config["categories"])
        for tid, cat in overrides.items():
            rec = self.ledger["transactions"].get(tid)
            if not rec or cat not in valid:
                continue
            baseline = rec.get("category_written") or rec.get("category_override") or rec.get("category")
            if cat == baseline:                        # unchanged since we last wrote the sheet
                continue
            if cat == rec.get("category"):            # set back to the automatic value -> drop override
                if rec.pop("category_override", None):
                    n += 1
            elif cat != rec.get("category_override"):
                rec["category_override"] = cat
                n += 1
        return n

    def recategorize(self, categorizer) -> int:
        """Re-run the current rules over every stored transaction (overrides are kept). Returns how many changed."""
        n = 0
        for rec in self.ledger["transactions"].values():
            new = categorizer.categorize(rec)
            if new != rec.get("category"):
                rec["category"] = new
                n += 1
        return n

    def visible_transactions(self) -> list[dict]:
        """Posted, non-removed transactions on non-hidden accounts, oldest first."""
        out = []
        for t in self.ledger["transactions"].values():
            if t.get("removed") or t.get("pending"):
                continue
            if t["account_id"] not in self.state["accounts"] or self.account_hidden(t["account_id"]):
                continue
            out.append(t)
        out.sort(key=lambda t: (t["date"] or "", t.get("posted_date") or "", t["id"]))
        return out

    # ---------------------------------------------------------------- balances
    def compute_balances(self, txns: list[dict]) -> None:
        """
        Plaid does not supply a per-transaction running balance, so we anchor on the
        bank-reported current balance and walk backwards through posted transactions.
        Sets t['balance_before'] and t['balance_after'] in place.

        Sign convention: depository -> balance_after = balance_before - amount
                         credit     -> owed_after    = owed_before    + amount
        """
        by_acct: dict[str, list[dict]] = {}
        for t in txns:
            by_acct.setdefault(t["account_id"], []).append(t)
        for aid, rows in by_acct.items():
            acct = self.state["accounts"].get(aid, {})
            current = acct.get("balance_current")
            if current is None:
                for t in rows:
                    t["balance_before"] = t["balance_after"] = None
                continue
            is_credit = acct.get("type") == "credit"
            running = float(current)
            for t in reversed(rows):           # newest -> oldest
                t["balance_after"] = round(running, 2)
                delta = t["amount"] if is_credit else -t["amount"]
                running = running - delta
                t["balance_before"] = round(running, 2)


def today_iso() -> str:
    return date.today().isoformat()
