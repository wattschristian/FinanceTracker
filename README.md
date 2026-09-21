# Household Finance Sync

Pulls transactions and balances from your banks through [Plaid](https://plaid.com) and rebuilds an Excel workbook
that tracks household spending across two people and any number of accounts. Built for a two-person household
where money moves between accounts constantly and neither a single budget app nor a hand-kept spreadsheet
kept up.

Run it, open the workbook. That is the whole routine.

Network access is limited to Plaid's API. Nothing is uploaded anywhere else, and the workbook, ledger and
credentials never leave the folder you run it from.

## What it produces

| Tab | What's on it |
|---|---|
| **Budget** | Your own planner sheet, untouched. Feeds target figures into the generated tabs. |
| **Transactions** | Every posted transaction, grouped by month then account, chronological by charge date. Account headers colored by owner, category cells colored by category and editable as dropdowns, running balance before/after each charge, per-account subtotals. |
| **Monthly Stats** | A section per person plus a household section: income, other inflows, spending, net, savings rate, top category, transaction count; averages and totals; a category × month matrix; budget variance, a savings projection and charts on the household section. Each section has an advice column that survives rebuilds. |
| **Category Budgets** | Derived spending caps and a monthly over/under verdict. Take-home income − that person's share of fixed commitments − savings target = discretionary pool, split across flexible categories by guideline percentages of take-home with dollar floors, scaled to fit. A part-month is judged against a pro-rated pace target. |
| **Accounts** | Linked accounts, owner, label, balances. Editable. |
| **Sync Log** | One line per run with counts and warnings. |

Three inflow buckets keep the income figure honest: **Income** is recurring pay, **Other Inflow** is one-off money
in (loan disbursements, aid, tax refunds), **Refund** is merchant refunds. Transfers between your own accounts are
excluded from every statistic — that alone fixed a five-figure phantom "income" month during development.

Pending transactions are held back until they post, so amounts never shift under you.

## Requirements

- Python 3.11+
- A Plaid account. The free Trial plan uses production data and allows up to 10 linked bank logins —
  see [Setting up Plaid](#setting-up-plaid) before you start linking.
- Excel, LibreOffice, or anything else that opens `.xlsx`.

## Quick start

```bash
git clone (https://github.com/wattschristian/FinanceTracker.git) finance-sync
cd finance-sync
./run.sh --demo          # Windows: run.bat --demo
```

`--demo` builds a workbook from invented data with no network access, so you can see the layout before linking
anything. When you're ready for real accounts:

```bash
./run.sh
```

First run creates a virtualenv, copies `config.example.json` to `config.json`, asks for the two household names
and your Plaid keys, then opens Plaid's hosted login page for each bank. Read
[Setting up Plaid](#setting-up-plaid) first — linking a bank spends one of a fixed number of free connections.
Full walkthrough in [SETUP.md](SETUP.md).

```
./run.sh              sync and open the workbook
./run.sh --link       link another bank
./run.sh --rebuild    re-render from the local ledger, no Plaid calls
./run.sh --demo       offline demo data
./run.sh --no-open    skip launching Excel
./run.sh --forget ID  unlink a bank   (does NOT refund the Item)
```

## Setting up Plaid

Plaid is the layer that talks to your banks. You need a free account and two keys; the program asks for them on
first run and writes them to `.env`.

### 1. Create a Plaid account

Sign up at [dashboard.plaid.com/signup](https://dashboard.plaid.com/signup) and verify your email, then apply for
a **Trial plan** from the Dashboard homepage (or at `dashboard.plaid.com/trial-plan`). The Trial plan is free and
lets you use real production data without a business registration, security questionnaire, or sales call. Most
applications are approved automatically after identity verification; if yours is flagged for manual review,
Plaid's team follows up by email within a few business days.

The Trial plan is only offered to teams that have never applied for Production access, so apply for it *before*
clicking anything that requests full Production.

### 2. Copy your keys

Once approved, open **Developers → Keys** in the Dashboard. You need two values:

| Value | Where it's shown |
|---|---|
| `client_id` | One per team, same across environments |
| `secret` | One per environment — take the **Production** secret for real banks, or **Sandbox** for testing |

### 3. Get them into the project

**Keys go in `.env`, not `config.json`.** They are credentials, not settings, and `.env` is the file both
`.gitignore` and the pre-commit hook guard most aggressively. `config.json` holds preferences — owners,
categories, rules, budget model — and never contains a secret.

Easiest path: run the program and paste when prompted.

```bash
./run.sh
```

It writes `.env` for you. Or create it yourself from the template:

```bash
cp .env.example .env
```

```ini
PLAID_CLIENT_ID=your_client_id_here
PLAID_SECRET=your_production_or_sandbox_secret_here
PLAID_ENV=production        # or: sandbox
```

To re-enter keys later, delete `.env` and run again.

### 4. Link your banks

With keys in place, `./run.sh` opens Plaid's hosted login page in your browser. Pick the institution, sign in on
the bank's own page, choose which accounts to share, and return to the console — it detects completion by
polling. It then asks who owns each account and pulls history.

> **⚠️ Every bank you link permanently spends one of your 10 free Items.**
>
> An **Item** is one login at one institution. The Trial plan caps you at 10 for the lifetime of the account, and
> **removing an Item with `--forget` does not give the slot back** — Plaid counts every access token ever created.
> Once all 10 are gone, the only way to link another bank is a paid plan.
>
> A typical two-person household uses 2–3: one login each, plus a second institution. **So get it right the first
> couple of times.** Before you link, know which login covers which accounts — one login that exposes four
> accounts costs one Item, while linking the same bank twice to catch a missed account costs two.
>
> Re-authenticating an *existing* bank is free. When a bank's consent lapses the tool reopens Plaid in update
> mode, which reuses the same Item and consumes nothing. The expensive mistake is turning a re-auth into a fresh
> link — which is exactly what happens if you run the program from a folder without your `state.json`, since it
> sees no linked banks and offers to link from scratch. **Never start over in a new folder; copy `state.json`,
> `ledger.json` and `.env` across instead.**

### Testing without spending Items

Sandbox is free, unlimited, and doesn't touch the Trial cap. Put your Sandbox secret in `.env` with
`PLAID_ENV=sandbox` and log into Plaid's test bank with `user_good` / `pass_good`. Switch to real banks by
replacing the secret, setting `PLAID_ENV=production`, and deleting `state.json` so the sandbox Items don't
linger.

`./run.sh --demo` is cheaper still: it fabricates a workbook offline with no Plaid account at all.

### OAuth timing

Trial plans include OAuth access to most major institutions — Bank of America, Chase, Wells Fargo, Capital One,
Citi, PNC, U.S. Bank, Amex and others. Access typically becomes available **6–24 hours after your account is
approved**, sometimes longer for a given bank. If a big bank refuses to connect on day one, that's usually why;
wait a day before assuming something is broken.

## Layout

```
finance_sync.py       entry point: setup wizard, linking, sync, rebuild
plaid_api.py          the only file that touches the network
ledger.py             local storage and running-balance math
categorize.py         merchant rules, then Plaid's category
workbook.py           builds the tabs, preserves your edits
demo_data.py          invented data for --demo
config.example.json   template; copied to config.json on first run
hooks/pre-commit      blocks commits containing private files
```

[ARCHITECTURE.md](ARCHITECTURE.md) explains how the modules fit together and why the workbook is rebuilt rather
than appended to.

## What stays out of git

`.gitignore` excludes the four things that carry personal data:

- `.env` — Plaid client id and secret
- `state.json` — linked items, cursors, balances, and access tokens when the OS keychain isn't available
- `ledger.json` — every transaction ever pulled
- `config.json` — Plaid account ids, real names, account labels, goal amounts, and merchant patterns naming
  your employer, landlord and utilities
- `*.xlsx` and `backups/` — the workbooks themselves

Only `config.example.json` and `.env.example` are committed. `demo_data.py` is entirely invented — no real
person, bank, employer or merchant appears in it. Keep it that way; mirror your own accounts in a local copy if
you want to.

## Publishing checklist

A `.gitignore` only stops future commits. Before making the repo public:

```bash
# 1. Is anything private already tracked?
git ls-files | grep -E '\.env$|config\.json|state\.json|ledger\.json|\.xlsx$|backups/'

# 2. Has it ever been committed, even if deleted since?
git log --all --name-only --pretty=format: | sort -u | \
  grep -E '\.env$|config\.json|state\.json|ledger\.json|\.xlsx$'
```

If either returns nothing, you're clear.

If the first returns files but the second doesn't, they're only staged or in the working tree:

```bash
git rm --cached <file>
git commit -m "Stop tracking private data"
```

If the second returns files, the data is in history and a later deletion does **not** remove it. Rewrite history
before pushing — [`git-filter-repo`](https://github.com/newren/git-filter-repo) is the maintained tool:

```bash
git filter-repo --invert-paths --path config.json --path state.json --path ledger.json --path .env
```

Then **rotate your Plaid secret** in the Plaid dashboard if `.env` was ever committed — a pushed secret should be
treated as compromised even if the push was to a private repo. If the repo has already been pushed public with
bank data in it, rotate first and rewrite second; deleting the repo does not reliably remove forks or caches.

Install the guard so this can't happen by accident:

```bash
git config core.hooksPath hooks
```

## Security notes

- Bank passwords are never seen by this program. Login happens on the bank's own OAuth page inside Plaid.
- Plaid access tokens go to the OS credential store (Windows Credential Manager / macOS Keychain / Secret
  Service) via `keyring`, falling back to `state.json` when that's unavailable.
- To revoke everything: `--forget <item id>` for each bank (calls Plaid's `/item/remove`), delete `.env`, delete
  the folder.

## License

No license has been chosen yet. Without one, default copyright applies and others may not reuse the code —
add a `LICENSE` file if you want that to change.
