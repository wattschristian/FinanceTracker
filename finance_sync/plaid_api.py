"""
Minimal Plaid client. Only talks to Plaid's API; nothing else on the network.

Uses plain HTTPS calls (requests) rather than the plaid-python SDK so there is
nothing to break when the SDK changes major versions. Endpoints used:

  /link/token/create        -> Hosted Link URL (Plaid hosts the bank-login page)
  /link/token/get           -> poll for the public_token after the user finishes
  /item/public_token/exchange
  /item/get + /institutions/get_by_id
  /accounts/get             -> balances
  /transactions/sync        -> incremental transactions (cursor based)
"""
from __future__ import annotations

import time
import webbrowser
from dataclasses import dataclass

import requests

BASE_URLS = {
    "sandbox": "https://sandbox.plaid.com",
    "production": "https://production.plaid.com",
}


class PlaidError(Exception):
    def __init__(self, code: str, message: str, request_id: str = ""):
        super().__init__(f"{code}: {message} (request_id={request_id})")
        self.code = code
        self.message = message
        self.request_id = request_id


class LinkAbandoned(Exception):
    """The user closed the Hosted Link page without finishing."""


@dataclass
class SyncResult:
    added: list
    modified: list
    removed: list
    next_cursor: str


class PlaidClient:
    def __init__(self, client_id: str, secret: str, env: str = "production", timeout: int = 60):
        if env not in BASE_URLS:
            raise ValueError(f"PLAID_ENV must be one of {list(BASE_URLS)}, got {env!r}")
        self.client_id = client_id
        self.secret = secret
        self.env = env
        self.base = BASE_URLS[env]
        self.timeout = timeout
        self._session = requests.Session()

    # ------------------------------------------------------------------ core
    def _post(self, path: str, body: dict) -> dict:
        payload = {"client_id": self.client_id, "secret": self.secret, **body}
        resp = self._session.post(self.base + path, json=payload, timeout=self.timeout)
        try:
            data = resp.json()
        except ValueError:
            raise PlaidError("HTTP_ERROR", f"Non-JSON response {resp.status_code}: {resp.text[:200]}")
        if resp.status_code >= 400 or data.get("error_code"):
            raise PlaidError(
                data.get("error_code", f"HTTP_{resp.status_code}"),
                data.get("error_message", resp.text[:200]),
                data.get("request_id", ""),
            )
        return data

    # ------------------------------------------------------------ hosted link
    def create_hosted_link(self, client_name: str, days_requested: int = 730,
                           access_token: str | None = None, lifetime_seconds: int = 1800) -> tuple[str, str]:
        """Returns (link_token, hosted_link_url). Pass access_token for update/re-auth mode."""
        body = {
            "client_name": client_name,
            "language": "en",
            "country_codes": ["US"],
            "user": {"client_user_id": "household"},
            "hosted_link": {"url_lifetime_seconds": lifetime_seconds},
        }
        if access_token:
            body["access_token"] = access_token          # update mode: no products list
        else:
            body["products"] = ["transactions"]
            body["transactions"] = {"days_requested": max(1, min(int(days_requested), 730))}
        data = self._post("/link/token/create", body)
        url = data.get("hosted_link_url")
        if not url:
            raise PlaidError("NO_HOSTED_LINK", "Plaid did not return a hosted_link_url; check that Hosted Link is available for your team.")
        return data["link_token"], url

    def wait_for_link_completion(self, link_token: str, timeout_seconds: int = 1800, poll_seconds: int = 3) -> str | None:
        """
        Poll /link/token/get until the user finishes the hosted Link session.
        Returns the public_token (None in update mode where no new token is issued).
        Raises LinkAbandoned if the user exits without linking.
        """
        deadline = time.time() + timeout_seconds
        while time.time() < deadline:
            data = self._post("/link/token/get", {"link_token": link_token})
            for session in data.get("link_sessions", []) or []:
                results = session.get("results") or {}
                for item in results.get("item_add_results", []) or []:
                    if item.get("public_token"):
                        return item["public_token"]
                on_success = session.get("on_success") or {}
                if on_success.get("public_token"):
                    return on_success["public_token"]
                if session.get("finished_at"):
                    # finished but no token -> either update mode success or user exit
                    on_exit = session.get("on_exit") or {}
                    err = (on_exit.get("error") or {}) if on_exit else {}
                    if err.get("error_code"):
                        raise LinkAbandoned(f"{err.get('error_code')}: {err.get('error_message')}")
                    if on_exit:
                        raise LinkAbandoned("Link window was closed before finishing.")
                    return None
            time.sleep(poll_seconds)
        raise LinkAbandoned("Timed out waiting for the Link session to finish.")

    def open_in_browser(self, url: str) -> None:
        print(f"\nOpening Plaid Link in your browser. If it does not open, paste this URL:\n  {url}\n")
        try:
            webbrowser.open(url)
        except Exception:
            pass

    # ------------------------------------------------------------------ items
    def exchange_public_token(self, public_token: str) -> tuple[str, str]:
        data = self._post("/item/public_token/exchange", {"public_token": public_token})
        return data["access_token"], data["item_id"]

    def institution_name(self, access_token: str) -> str:
        item = self._post("/item/get", {"access_token": access_token})["item"]
        ins_id = item.get("institution_id")
        if not ins_id:
            return "Unknown institution"
        try:
            ins = self._post("/institutions/get_by_id", {"institution_id": ins_id, "country_codes": ["US"]})
            return ins["institution"]["name"]
        except PlaidError:
            return ins_id

    def remove_item(self, access_token: str) -> None:
        self._post("/item/remove", {"access_token": access_token})

    # --------------------------------------------------------------- accounts
    def get_accounts(self, access_token: str) -> list[dict]:
        return self._post("/accounts/get", {"access_token": access_token})["accounts"]

    # ----------------------------------------------------------- transactions
    def transactions_sync(self, access_token: str, cursor: str | None) -> SyncResult:
        """Fetch every page since `cursor`. Restarts from the original cursor if Plaid reports a mutation mid-pagination."""
        for attempt in range(3):
            added, modified, removed = [], [], []
            page_cursor = cursor
            try:
                while True:
                    body = {"access_token": access_token, "count": 500,
                            "options": {"include_personal_finance_category": True, "include_original_description": True}}
                    if page_cursor:
                        body["cursor"] = page_cursor
                    data = self._post("/transactions/sync", body)
                    added += data.get("added", [])
                    modified += data.get("modified", [])
                    removed += data.get("removed", [])
                    page_cursor = data["next_cursor"]
                    if not data.get("has_more"):
                        return SyncResult(added, modified, removed, page_cursor)
            except PlaidError as e:
                if e.code == "TRANSACTIONS_SYNC_MUTATION_DURING_PAGINATION" and attempt < 2:
                    time.sleep(2)
                    continue
                raise
        raise PlaidError("SYNC_RETRY_EXHAUSTED", "Could not complete transactions sync after 3 attempts.")
