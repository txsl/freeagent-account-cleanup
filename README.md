# FreeAgent Revolut cleanup — setup & workflow

## 1. Install

Use your virtual environment (e.g., `virtualenv`) of choice, then install:
```bash
pip install -r requirements.txt
```

## 2. Register an app and get a refresh token (one-off)

FreeAgent's API is OAuth 2.0 only — there's no personal-access-token option.
`get_refresh_token.py` does the one-off authorization locally, so your
Client Secret never passes through a third party.

1. Register an app at the [Developer Dashboard](https://dev.freeagent.com/apps)
   (production, not sandbox). Note the **Client ID** and **Client Secret**.
2. On the same app, add a redirect URI: `http://localhost:53682/callback`
   (must match exactly — this is what the script listens on).
3. `cp .env.example .env` and fill in `FREEAGENT_CLIENT_ID` /
   `FREEAGENT_CLIENT_SECRET`. Leave `FREEAGENT_REFRESH_TOKEN` blank.
4. Run:
   ```bash
   python get_refresh_token.py
   ```
   This opens your browser to FreeAgent's approve screen, catches the
   redirect on a local server, exchanges the code for tokens immediately
   (authorization codes expire after 15 minutes), and writes
   `FREEAGENT_REFRESH_TOKEN` into `.env` for you.

You won't need to repeat this — refresh tokens are long-lived (~20 years)
and the other scripts renew the access token from it automatically each run.

## 3. Configure

Confirm `.env` has `FREEAGENT_CLIENT_ID`, `FREEAGENT_CLIENT_SECRET`, and
`FREEAGENT_REFRESH_TOKEN` all filled in (the last one comes from step 2
above).

Then:
```bash
python list_accounts.py
```
This prints every bank account with its ID. Copy the polluted account's ID
and the new clean account's ID into `.env` as `ACCOUNT_POLLUTED_ID` /
`ACCOUNT_CLEAN_ID`.

## 4. Build the review workbook

```bash
python build_review_workbook.py
```

This fetches both accounts (caching the raw "before" state to `cache/` for
your audit trail), and writes `review_<timestamp>.xlsx` with:

- **Instructions** — what everything means, read this first
- **Account A (Polluted)** — raw transactions, plus `match_count` /
  `match_status` (plain `COUNTIFS`/`IF` formulas — nothing fancier) and a
  blank `approve_delete` column for you to fill in
- **Account B (Clean)** — raw transactions, unedited

The console output also prints each account's earliest transaction date —
check the clean account's earliest date against how far back the
misconfigured feed actually ran. Anything older than that falls outside
what this diff can safely catch.

## 5. Review

Open the workbook. Filter Account A to `match_status = matched`, spot-check
a handful against Account B, and mark `approve_delete = Y` on the rows
you're confident about. For anything `AMBIGUOUS - review` (which implies
multiple line with the same date & transaction in either Account A or B),
filter Account A and B by that date to see the candidates yourself. Pay attention
to `explanation_type`, `is_deletable`, and `is_locked` before approving anything
that's explained — the Instructions tab spells out what each means.

Save the file.

## 6. Delete (dry run first, always)

```bash
python delete_transactions.py review_<timestamp>.xlsx
```

This only *simulates* and logs to `delete_log_<timestamp>.csv` — nothing is
changed. Check the log matches what you expect, then:

```bash
python delete_transactions.py review_<timestamp>.xlsx --live
```

You'll be asked to type `DELETE` to confirm. It re-checks each transaction's
live state right before acting (in case anything changed since you built
the workbook), skips anything not currently deletable or flagged as an
Invoice Receipt, and logs every action.

**One thing worth testing carefully the first time:** FreeAgent's own docs
show the transaction-delete endpoint as `/v2/bank_transaction/:id`
(singular — not `/bank_transactions/:id` like everything else). Other
integrators have reported confusing errors here. Consider approving just
one low-stakes row first, running `--live` on that alone, and confirming
in the FreeAgent UI that it worked as expected before doing the rest.

## Running the tests

The test suite mocks every FreeAgent API call — no credentials or network
access needed — and covers token refresh, pagination, the singular
`/bank_transaction/:id` delete endpoint, the matching formulas (run through
an actual LibreOffice recalculation, not just checked as text), and every
branch of the delete logic including the live-run confirmation gate.

```bash
pip install -r requirements-dev.txt
pytest
```

## Notes

- Deletion is irreversible. The dry run and the `cache/` snapshot are your
  safety net — keep both until you're satisfied.
- Rate limits are 120 requests/minute and 3600/hour; the scripts pace
  themselves well under that, but a very large account may still take a
  few minutes to fetch.
