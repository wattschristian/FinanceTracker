# Household Finance Sync — Setup Guide

A small Python tool that pulls transactions and balances from your banks (through Plaid) and rebuilds
`Financial_Tracking.xlsx`. After setup the monthly routine is: **double-click `run.bat`, open the spreadsheet.**
It works with any institution Plaid supports; it has been used against a bank + credit-union pair with
OAuth logins.

The only network traffic the program makes is to `production.plaid.com` (or `sandbox.plaid.com`). It never
uploads the spreadsheet or your ledger anywhere.

---

## What's in the folder

| File | Purpose |
|---|---|
| `run.bat` / `run.sh` | Double-click launcher. Creates a private Python environment on first run, installs the four dependencies, runs the sync. |
| `finance_sync.py` | Entry point: setup wizard, bank linking, sync, workbook rebuild. |
| `plaid_api.py` | The only file that touches the network. Plain HTTPS calls to Plaid. |
| `ledger.py` | Local storage (`state.json`, `ledger.json`) and running-balance math. |
| `categorize.py` | Assigns spending categories (your merchant rules first, then Plaid's category). |
| `workbook.py` | Builds the Excel tabs and preserves anything you typed into them. |
| `demo_data.py` | Fake data for `--demo`. |
| `config.example.json` | The template that ships with the repo. Copied to `config.json` on first run. |
| `config.json` | **Yours to edit**, never committed: owners and colors, categories and colors, merchant rules, which Budget cells feed which category, savings goal, account owners. |
| `Financial_Tracking.xlsx` | Your workbook. The original planner sheet becomes the `Budget` tab and is never modified. |
| `.env` | Created on first run: your Plaid keys. Keep private. |
| `state.json` / `ledger.json` | Created on first run: linked banks, cursors, balances, every transaction received. Keep private. |
| `backups/` | The previous 10 versions of the workbook, one per run. |

---

## Step 1 — Install Python (once)

Windows: download Python 3.11 or newer from <https://www.python.org/downloads/>, run the installer, and tick
**"Add python.exe to PATH"** on the first screen. macOS/Linux: `python3 --version` should be 3.11+.

## Step 2 — Put the folder somewhere private

Unzip `finance_sync` into a folder only you can read, e.g. `C:\Users\<you>\Documents\FinanceSync`. Not a shared
drive, not OneDrive/Dropbox unless you are comfortable with those services seeing `state.json` and `ledger.json`.

## Step 3 — Create a Plaid account and get your keys (once)

1. Go to <https://dashboard.plaid.com/signup> and sign up. Choose the **Trial plan** when offered — it is free,
   uses real production data, allows up to 10 linked bank logins ("Items"), and includes OAuth access to Bank of
   America without a separate production application. Most applications are approved automatically after an identity check;
   if yours is flagged for manual review Plaid follows up within a few business days.
2. Once approved, open **Developers → Keys**. Copy the `client_id` and the **Production** `secret`.
3. Plaid says OAuth access to the big banks usually turns on 6–24 hours after approval. If a major bank refuses
   to connect on day one, wait a day and try again.

Important about the 10-Item limit: an Item is one bank login. Removing an Item does **not** give the slot back on
the Trial plan, so don't link and unlink for fun. You'll likely use 2–3 slots total:

* Person 1's bank login (1 Item)
* Person 2's login at the same bank, if it is a separate login (1 Item)
* The second institution (1 Item)

If you want to try the tool with fake banks first, use the **Sandbox** secret instead and answer `sandbox` when
asked for `PLAID_ENV`. In Sandbox the test bank accepts username `user_good` / password `pass_good`. Sandbox
does not count against the 10 Items. To switch to real banks later, delete `.env` and `state.json` and run again.

## Step 4 — First run

Double-click `run.bat` (or `./run.sh`). It will:

1. Create `config.json` from `config.example.json` and ask for the two household names.
2. Ask for `PLAID_CLIENT_ID`, `PLAID_SECRET`, `PLAID_ENV` and save them to `.env`.
3. Open a Plaid page in your browser. Pick your bank, sign in, choose the accounts to share, finish. Come back to
   the console — it notices automatically. It then asks **"Link another bank?"** — say `y` and repeat for each login.
4. Ask who owns each account (by number, from the names you gave) and what label to show. Answers go into
   `config.json → accounts`; you can change them later there or on the **Accounts** tab of the workbook.
5. Pull up to 24 months of history (limited by `history_start_date` in `config.json`, default 2026‑08‑01),
   categorize it, rebuild the workbook, and open it.

The first pull can take a minute per bank; Plaid sometimes needs a moment to prepare history and the tool retries.

## Step 5 — The monthly routine

1. Double-click `run.bat`. Close Excel first if the workbook is open (the tool will tell you if it isn't).
2. Open `Financial_Tracking.xlsx`.
3. Send me the workbook. I read the **Monthly Stats** tab and write the *Advice for next month* column. Paste my
   text into that column — it is preserved on every future sync.

That's it. Everything else is automatic.

---

## The workbook

**Budget** — your original sheet, untouched. `B15` (total monthly expenses), `C15` (planned income) and `D15`
(left over) feed the stats; the per-line figures feed the category budget column via `config.json →
budget_sheet.category_cells`. If you add rows to the budget, update those cell references.

**Transactions** — every posted transaction, grouped by month, then by account, oldest first within a group.
Account group headers are colored by owner; the Category cell is colored by category and is a dropdown — change it
and the next sync keeps your choice for that transaction. Balance Before / After are computed by walking backwards
from the balance the bank reported at sync time (Plaid does not supply per-transaction balances). Transfers between
your own accounts show in both accounts and are excluded from income/spending statistics.

**Monthly Stats** — one section per owner (plus Joint when joint accounts have activity) followed by a
**Household** section. Each section has a month table (income, other inflows, spending, net, savings rate, top
category, transaction count; the household table adds budget variance), averages and totals, and a category × month
matrix. The household section adds the budget comparison, the savings projection (goal amount and target date are
blue input cells) and two charts. Every section has its own *Advice for next month* column that survives rebuilds.

Three inflow buckets keep the income figure honest: **Income** is recurring pay and interest; **Other Inflow** is
one-off money in (loan disbursements, financial aid, tax refunds) — tracked in its own column and in "Net incl.
other inflows" but kept out of Income and the savings rate; **Refund** is merchant refunds. Transfers between your own
accounts (vault moves, Zelle/Cash App between you, card payments) are excluded from every statistic.

**Category Budgets** — a derived spending plan and a monthly over/under verdict, one section per owner plus a
household section. Nothing here is typed in by hand:

    take-home income (recurring pay, averaged over complete months)
      − that person's share of household fixed commitments
      − the savings target
      = discretionary pool

The pool is divided across the flexible categories using guideline percentages of take-home pay (groceries 10% with
a USDA moderate-cost floor of ~$785/mo for two adults, dining 5%, transportation 8%, and so on), floored at the
dollar minimums, and scaled down proportionally when the guidelines add up to more than the pool. If even the
floors exceed the pool, a warning line says so and every cap is scaled to what is actually available.

Fixed bills are charged by *share*, not by whose account paid them, so one of you covering rent while the other
transfers money doesn't wreck either budget. The **Share of shared costs** cell defaults to that person's share of
household income — change it if one of you does all the grocery shopping or pays more of the bills.

The month in focus is a dropdown at the top. A part-month is judged against a pro-rated "pace target" (cap × share
of the month elapsed) so an in-progress month doesn't look artificially good, and the Over/(Under) column is
colored red or green. The blue cells — month, savings rate, every guideline % and floor, each split — are yours to
change and are kept across syncs.

**Accounts** — linked accounts with owner, label, balances. Edit Owner/Label/Hidden here; the next sync applies it.

**Sync Log** — one line per run with counts and any warnings.

Pending transactions are held back until they post, so the sheet never shows a charge whose amount might change.

---

## Tuning categories

`config.json → merchant_rules` is a list of regular expressions matched (case-insensitive) against the merchant
name and description; the first match wins and beats Plaid's own category. Add a line for anything that keeps
landing in the wrong bucket, e.g. `{ "pattern": "local coffee co", "category": "Dining" }`. To rename or recolor a
category edit `categories` (colors are hex without `#`). A category used in a rule must exist in `categories`.

One-off fixes are easier in the sheet: change the Category cell and run the sync — it becomes an override for that
transaction. Setting a cell back to the automatic value removes the override.

---

## Updating the tool

Pull the new `.py` files. Your `config.json` is not tracked, so it survives a `git pull` untouched — if a release
adds config keys, diff it against `config.example.json` and copy the new keys across. Rule changes apply to history too: every run re-categorizes the whole ledger with the current rules
(manual overrides in the sheet are kept), so a corrected rule fixes old months on the next sync or `--rebuild`.

## Troubleshooting

| Symptom | Fix |
|---|---|
| "…is open in Excel. Close it and run again." | Close the workbook and rerun. |
| Plaid error `INVALID_API_KEYS` | Wrong secret for the environment (Sandbox vs Production). Delete `.env` and rerun. |
| A bank asks you to sign in again | Normal for OAuth banks every few months. The tool opens Plaid for re-authentication automatically; finish it and the sync continues. |
| `ITEM_LOGIN_REQUIRED` keeps coming back | Run `run.bat --forget <item id>` (ids are on the Accounts tab) and link the bank again. This uses one of the 10 Trial slots. |
| Balance After on the newest row doesn't match the bank app | Usually pending transactions the bank already counts. It corrects itself once they post. |
| A transaction is in the wrong category | Change the Category cell (one-off) or add a `merchant_rules` entry (permanent). |
| A big bank won't connect right after signup | OAuth access is enabled 6–24 h after Plaid approval; try again tomorrow. |
| Want to see the layout before linking anything | `run.bat --demo` builds `Financial_Tracking_DEMO.xlsx` from fake data, offline. |
| Want to re-render without calling Plaid | `run.bat --rebuild`. |

Command-line flags: `--link` (add a bank), `--demo`, `--rebuild`, `--no-open`, `--forget ITEM_ID`.

## Automating it (optional)

Windows Task Scheduler → Create Basic Task → Monthly → Action: *Start a program* → Program:
`C:\Users\<you>\Documents\FinanceSync\run.bat`, Arguments: `--no-open`. If a bank needs re-authentication the run
will wait for you in the browser, so a scheduled run works best on a day you're at the computer.

## Security notes

* Your bank passwords are never seen by this program; the login happens on the bank's own OAuth page inside Plaid.
* Plaid access tokens are stored in the Windows Credential Manager / macOS Keychain (via `keyring`). If that is
  unavailable they fall back to `state.json`.
* `.env`, `state.json`, `ledger.json`, the workbook and `backups/` contain your financial data. Keep the folder
  private and out of any repo (`.gitignore` is included).
* To revoke everything: `run.bat --forget <item id>` for each bank (this calls Plaid's `/item/remove`), delete
  `.env`, and delete the folder.
