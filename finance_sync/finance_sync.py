#!/usr/bin/env python3
"""
Household Finance Sync
======================
Pulls transactions + balances from your banks through Plaid and rebuilds the
tracking workbook. Network access is limited to Plaid's API.

Usage (or just double-click run.bat / run.sh):
  python finance_sync.py            sync everything and open the workbook
  python finance_sync.py --link     link another bank, then sync
  python finance_sync.py --demo     build Financial_Tracking_DEMO.xlsx from fake data (no network)
  python finance_sync.py --no-open  don't launch Excel afterwards
  python finance_sync.py --rebuild  re-render the workbook from the local ledger without calling Plaid
  python finance_sync.py --forget ITEM_ID   unlink a bank (removes it at Plaid and locally)
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
os.chdir(HERE)

try:
    from dotenv import load_dotenv
except ImportError:  # pragma: no cover
    load_dotenv = None

from categorize import Categorizer
from ledger import Store
from plaid_api import LinkAbandoned, PlaidClient, PlaidError
import workbook as wbmod


# ------------------------------------------------------------------ helpers
def say(msg: str = "") -> None:
    print(msg, flush=True)


def ask(prompt: str, default: str | None = None) -> str:
    suffix = f" [{default}]" if default else ""
    val = input(f"{prompt}{suffix}: ").strip()
    return val or (default or "")


def yes(prompt: str, default: bool = True) -> bool:
    d = "Y/n" if default else "y/N"
    v = input(f"{prompt} ({d}): ").strip().lower()
    if not v:
        return default
    return v in ("y", "yes")


def open_file(path: Path) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as e:  # pragma: no cover
        say(f"(Could not auto-open the workbook: {e})")


# --------------------------------------------------------- config bootstrap
def bootstrap_config() -> None:
    """A fresh clone has config.example.json but no config.json (it is gitignored). Create it, then
    ask for the household's names so the placeholder owners don't end up in the spreadsheet."""
    cfg_path, example = HERE / "config.json", HERE / "config.example.json"
    if cfg_path.exists():
        return
    if not example.exists():
        raise FileNotFoundError("Neither config.json nor config.example.json is present in this folder.")
    shutil.copy2(example, cfg_path)
    say("\n== First-time setup: who lives in this household? ==")
    say("These names label the account groups and the per-person sections in the workbook.")
    cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
    placeholders = [o for o in cfg["owners"] if o.lower().startswith("owner ")]
    renamed = {}
    for i, old in enumerate(placeholders, start=1):
        new = ask(f"Name of person {i}", old).strip() or old
        renamed[old] = new
    if renamed:
        cfg["owners"] = {renamed.get(k, k): v for k, v in cfg["owners"].items()}
        cfg_path.write_text(json.dumps(cfg, indent=2) + "\n", encoding="utf-8")
    say(f"Created {cfg_path.name}. It stays out of git — edit it freely.\n")


# ------------------------------------------------------------ credentials
def load_credentials() -> tuple[str, str, str]:
    env_path = HERE / ".env"
    if load_dotenv and env_path.exists():
        load_dotenv(env_path)
    cid, sec, env = os.getenv("PLAID_CLIENT_ID"), os.getenv("PLAID_SECRET"), os.getenv("PLAID_ENV", "production")
    if cid and sec:
        return cid, sec, env.lower()

    say("\n== First-time setup: Plaid credentials ==")
    say("Find these at https://dashboard.plaid.com/developers/keys  (use the *Production* secret for your Trial plan;")
    say("use the Sandbox secret + PLAID_ENV=sandbox if you only want to test with fake banks).")
    cid = ask("PLAID_CLIENT_ID")
    sec = ask("PLAID_SECRET")
    env = ask("PLAID_ENV (production or sandbox)", "production").lower()
    env_path.write_text(f"PLAID_CLIENT_ID={cid}\nPLAID_SECRET={sec}\nPLAID_ENV={env}\n", encoding="utf-8")
    say(f"Saved to {env_path} (keep this file private).\n")
    return cid, sec, env


# ---------------------------------------------------------------- linking
def link_new_bank(client: PlaidClient, store: Store) -> str | None:
    cfg = store.config
    say("\nA Plaid page will open in your browser. Pick your bank, sign in, and choose the accounts to share.")
    say("Come back here when the page says you're done.")
    link_token, url = client.create_hosted_link(cfg.get("plaid_client_name", "Household Finance Sync"),
                                                days_requested=cfg.get("plaid_days_requested", 730))
    client.open_in_browser(url)
    try:
        public_token = client.wait_for_link_completion(link_token)
    except LinkAbandoned as e:
        say(f"Link not completed: {e}")
        return None
    if not public_token:
        say("Plaid returned no token; nothing linked.")
        return None
    access_token, item_id = client.exchange_public_token(public_token)
    store.state["items"][item_id] = {"institution": "…", "cursor": None, "linked_at": datetime.now().isoformat(timespec="minutes")}
    store.set_access_token(item_id, access_token)
    try:
        store.state["items"][item_id]["institution"] = client.institution_name(access_token)
    except PlaidError:
        store.state["items"][item_id]["institution"] = "Bank"
    store.save()
    say(f"Linked {store.state['items'][item_id]['institution']}  (item {item_id}).")
    return item_id


def relink_bank(client: PlaidClient, store: Store, item_id: str) -> bool:
    inst = store.state["items"][item_id].get("institution", "your bank")
    say(f"\n{inst} needs you to sign in again (this happens periodically with OAuth banks). Opening Plaid…")
    link_token, url = client.create_hosted_link(store.config.get("plaid_client_name", "Household Finance Sync"),
                                                access_token=store.get_access_token(item_id))
    client.open_in_browser(url)
    try:
        client.wait_for_link_completion(link_token)
        return True
    except LinkAbandoned as e:
        say(f"Re-authentication not completed: {e}")
        return False


def assign_owners(store: Store, new_ids: list[str]) -> None:
    owners = list(store.config["owners"])
    if not new_ids:
        return
    say("\n== New accounts found. Who does each belong to? ==")
    for i, o in enumerate(owners, start=1):
        say(f"  {i}) {o}")
    for aid in new_ids:
        a = store.state["accounts"][aid]
        default_label = store.account_label(aid)
        while True:
            choice = ask(f"{default_label}  [{a.get('type')}/{a.get('subtype')}] -> owner number", "1")
            if choice.isdigit() and 1 <= int(choice) <= len(owners):
                break
            say("  enter a number from the list")
        label = ask("  label to show in the spreadsheet", default_label)
        store.config["accounts"][aid] = {"owner": owners[int(choice) - 1], "label": label, "hidden": False}
    store.save()


# ------------------------------------------------------------------- sync
def sync_all(client: PlaidClient, store: Store, categorizer: Categorizer) -> dict:
    totals = {"added": 0, "modified": 0, "removed": 0, "pending": 0, "items": 0, "notes": []}
    now = datetime.now()
    for item_id in list(store.state["items"]):
        token = store.get_access_token(item_id)
        inst = store.state["items"][item_id].get("institution", item_id)
        if not token:
            totals["notes"].append(f"{inst}: no access token stored; run --forget {item_id} and link again.")
            continue
        for attempt in (1, 2):
            try:
                say(f"Syncing {inst}…")
                accounts = client.get_accounts(token)
                new_ids = store.upsert_accounts(item_id, accounts, now)
                assign_owners(store, new_ids)
                res = client.transactions_sync(token, store.state["items"][item_id].get("cursor"))
                counts = store.apply_sync(res.added, res.modified, res.removed, categorizer)
                store.state["items"][item_id]["cursor"] = res.next_cursor
                store.state["items"][item_id]["last_sync"] = now.isoformat(timespec="minutes")
                for k in ("added", "modified", "removed"):
                    totals[k] += counts[k]
                totals["pending"] = counts["pending_skipped"]
                totals["items"] += 1
                store.save()
                say(f"  {counts['added']} new, {counts['modified']} updated, {counts['removed']} removed")
                break
            except PlaidError as e:
                if e.code in ("ITEM_LOGIN_REQUIRED", "PENDING_EXPIRATION", "PENDING_DISCONNECT") and attempt == 1:
                    if relink_bank(client, store, item_id):
                        continue
                    totals["notes"].append(f"{inst}: needs re-authentication; run again and finish the Plaid page.")
                elif e.code == "PRODUCT_NOT_READY" and attempt == 1:
                    say("  Plaid is still preparing this bank's history; will retry in 15 s…")
                    import time; time.sleep(15)
                    continue
                else:
                    totals["notes"].append(f"{inst}: {e.code} – {e.message}")
                    say(f"  ! {e.code}: {e.message}")
                break
    store.state["last_sync"] = now.isoformat(timespec="minutes")
    store.save()
    return totals


def build_workbook(store: Store, categorizer: Categorizer, path: Path, mode: str, totals: dict) -> None:
    edits = wbmod.capture_user_edits(path)
    # category cells typed into the sheet become overrides
    changed = store.apply_category_overrides(edits["category_overrides"])
    # owner / label / hidden typed into the Accounts tab
    owners = set(store.config["owners"])
    for aid, e in edits["account_edits"].items():
        if aid in store.state["accounts"]:
            entry = store.config["accounts"].setdefault(aid, {})
            if e.get("owner") in owners:
                entry["owner"] = e["owner"]
            if e.get("label"):
                entry["label"] = e["label"]
            entry["hidden"] = bool(e.get("hidden"))
    recat = store.recategorize(categorizer)   # config.json rule changes apply to history too
    store.save()

    txns = store.visible_transactions()
    store.compute_balances(txns)
    notes = list(totals.get("notes", []))
    if changed:
        notes.append(f"{changed} category edit(s) from the sheet kept as overrides.")
    if recat:
        notes.append(f"{recat} transaction(s) re-categorized by updated rules.")
    log = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"), "mode": mode, "items": totals.get("items", 0),
        "added": totals.get("added", 0), "modified": totals.get("modified", 0), "removed": totals.get("removed", 0),
        "pending": totals.get("pending", 0), "notes": " | ".join(notes) if notes else "",
    }
    wbmod.render(store, txns, path, log, edits)
    store.save()   # persists category_written markers


# ------------------------------------------------------------------- demo
def run_demo(store: Store, categorizer: Categorizer, open_after: bool) -> None:
    import demo_data
    real = Path(store.config["workbook_path"])
    demo_path = real.with_name(real.stem + "_DEMO" + real.suffix)
    if real.exists() and not demo_path.exists():
        shutil.copy2(real, demo_path)     # keep the Budget tab so the stats formulas have something to point at
    # in-memory only: never touch the real state/ledger
    store.state = {"items": {k: {**v, "cursor": None} for k, v in demo_data.DEMO_ITEMS.items()}, "accounts": {}, "last_sync": None}
    store.ledger = {"transactions": {}}
    store.config = {**store.config, "accounts": {}}
    now = datetime.now()
    for item_id, accts in demo_data.demo_accounts_by_item().items():
        store.upsert_accounts(item_id, accts, now)
    # the demo ships generic owners; map them onto whatever names this household configured
    conf_owners = list(store.config["owners"])
    joint = next((o for o in conf_owners if o.strip().lower() == "joint"), None)
    personal = [o for o in conf_owners if o != joint] or conf_owners
    owner_map = {
        "Owner A": personal[0],
        "Owner B": personal[1] if len(personal) > 1 else personal[0],
        "Joint": joint or conf_owners[-1],
    }
    for a in demo_data.DEMO_ACCOUNTS:
        store.config["accounts"][a["account_id"]] = {
            "owner": owner_map.get(a["owner"], a["owner"]),
            "label": f"{demo_data.DEMO_ITEMS[a['item']]['institution']} {a['name']} ••{a['mask']}",
            "hidden": False,
        }
    start = date.fromisoformat(store.config.get("history_start_date", "2026-08-01"))
    txns = demo_data.demo_transactions(start, date.today())
    counts = store.apply_sync(txns, [], [], categorizer)
    store.save = lambda: None  # type: ignore[assignment]  # demo never persists
    totals = {"added": counts["added"], "modified": 0, "removed": 0, "pending": counts["pending_skipped"], "items": 2,
              "notes": ["DEMO DATA – nothing here is real."]}
    build_workbook(store, categorizer, demo_path, "demo", totals)
    say(f"\nDemo workbook written: {demo_path}")
    if open_after:
        open_file(demo_path)


# ------------------------------------------------------------------- main
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Sync bank transactions into the tracking workbook.")
    ap.add_argument("--link", action="store_true", help="link another bank before syncing")
    ap.add_argument("--demo", action="store_true", help="offline demo with fake data")
    ap.add_argument("--rebuild", action="store_true", help="re-render from local ledger, no Plaid calls")
    ap.add_argument("--no-open", action="store_true", help="do not open the workbook afterwards")
    ap.add_argument("--forget", metavar="ITEM_ID", help="unlink a bank")
    args = ap.parse_args(argv)

    try:
        bootstrap_config()
        store = Store(HERE)
        categorizer = Categorizer(store.config)
        open_after = store.config.get("open_workbook_after_sync", True) and not args.no_open
        wb_path = Path(store.config["workbook_path"])
        if not wb_path.is_absolute():
            wb_path = HERE / wb_path

        if args.demo:
            run_demo(store, categorizer, open_after)
            return 0

        if args.rebuild:
            build_workbook(store, categorizer, wb_path, "rebuild", {"items": 0})
            say(f"Workbook rebuilt: {wb_path}")
            if open_after:
                open_file(wb_path)
            return 0

        cid, sec, env = load_credentials()
        client = PlaidClient(cid, sec, env)
        if env == "sandbox":
            say("(Sandbox mode: use user_good / pass_good on the Plaid test bank.)")

        if args.forget:
            item = store.state["items"].pop(args.forget, None)
            if not item:
                say("No such item id. Ids are listed on the Accounts tab.")
                return 1
            tok = store.get_access_token(args.forget)
            if tok:
                try:
                    client.remove_item(tok)
                except PlaidError as e:
                    say(f"(Plaid could not remove it remotely: {e.code}; removed locally anyway)")
            for aid in [a for a, v in store.state["accounts"].items() if v.get("item_id") == args.forget]:
                store.state["accounts"].pop(aid, None)
            store.save()
            say(f"Forgot {item.get('institution', args.forget)}.")
            return 0

        if args.link or not store.state["items"]:
            if not store.state["items"]:
                say("\nNo banks linked yet. Let's link the first one.")
            while True:
                link_new_bank(client, store)
                if not yes("Link another bank?", default=False):
                    break

        totals = sync_all(client, store, categorizer)
        build_workbook(store, categorizer, wb_path, env, totals)
        say(f"\nDone. {totals['added']} new transaction(s). Workbook: {wb_path}")
        for n in totals["notes"]:
            say(f"  note: {n}")
        if open_after:
            open_file(wb_path)
        return 0

    except PermissionError as e:
        say(f"\n{e}")
        return 2
    except KeyboardInterrupt:
        say("\nCancelled.")
        return 130
    except PlaidError as e:
        say(f"\nPlaid error {e.code}: {e.message}")
        if e.code in ("INVALID_API_KEYS", "INVALID_INPUT"):
            say("Check PLAID_CLIENT_ID / PLAID_SECRET / PLAID_ENV in .env (delete .env to re-enter them).")
        return 3


if __name__ == "__main__":
    sys.exit(main())
