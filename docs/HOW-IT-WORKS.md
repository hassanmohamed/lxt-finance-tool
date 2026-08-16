# How It Works

For whoever has to change the code. Everything is in [app.py](../app.py) (~3,000 lines), in this
order: styling → settings → login → QuickBooks → transforms → pivot → UI → the ETL runner.

Streamlit re-runs the whole file on every click, so anything that must survive a click is kept in
`st.session_state`.

---

## The flow

```
login
  ↓
pick dates, upload mapping sheet + rate file
  ↓
_run_etl()                                          app.py:2888
  │
  └─ for each of the 9 companies:
        refresh_access_token()   get an access token
        _save_refresh_token()    save the new refresh token straight away
        fetch_general_ledger()   download the General Ledger
        flatten_report()         turn nested QuickBooks JSON into flat rows
        transform()              rename columns, work out currency and month
  ↓
combine all 9 into master_df
  ↓
apply_mapping()   add Mapping / Item / Statement from the mapping sheet
forex lookup      pick a rate from currency + month
                  → Amount in USD
  ↓
Excel export · Pivot P&L · dashboard · AI assistant
```

---

## Login

`check_password()` — [app.py:667](../app.py#L667)

One shared account. The password is checked with bcrypt against `APP_PASSWORD_HASH`. Five wrong
attempts locks you out for 15 minutes; 30 minutes of inactivity ends the session. Both are held in
memory, so restarting the app clears them.

---

## QuickBooks

One Intuit app is authorised against all 9 companies. Each company adds its own `realm_id` and
`refresh_token`.

- `refresh_access_token()` ([app.py:816](../app.py#L816)) — swaps a refresh token for an access token.
- `fetch_general_ledger()` ([app.py:893](../app.py#L893)) — downloads the report, accrual basis.
- `flatten_report()` ([app.py:924](../app.py#L924)) — QuickBooks returns a nested tree of rows inside
  rows; this walks it and produces one flat record per line.

**Column names are inconsistent.** QuickBooks sometimes returns internal keys (`debt_amt`) and
sometimes display names (`Debit`), so `QB_COLUMN_MAP` ([app.py:618](../app.py#L618)) handles both.
For multi-currency companies it also returns `Foreign Debit` / `Foreign Credit` — these are
deliberately **not** used, because the app wants home-currency amounts and converts them itself.
Using the foreign columns would convert twice.

---

## Tokens — the important part

QuickBooks refresh tokens are **single-use**. Every time you authenticate, the old token dies and a
new one is issued. Tokens also expire after 100 days unused.

The only place they are stored is `.streamlit/secrets.toml`. `_save_refresh_token()`
([app.py:852](../app.py#L852)) swaps the old token for the new one in the file text, which keeps
comments and formatting intact.

Two things here are deliberate:

- The save happens **inside the loop, right after each company authenticates**
  ([app.py:2939](../app.py#L2939)) — not at the end. If company 7 fails, companies 1–6 have already
  saved their new tokens.
- Every way the save can fail shows a warning in the app and logs an error. A silent failure would
  mean burning a token on every run without ever saving one, and the only way back is re-authorising
  each company through Intuit.

**`app.py` is the only thing that should ever authenticate with these credentials.** A separate
script or scheduled job with its own copy of the tokens will consume them without saving the
replacements. (A standalone CLI, a Render deployment, and a GitHub-gist token store all used to
exist here and were removed for this reason.) If you need scheduled extraction, build it inside this
app so it uses the same save path.

---

## Turning rows into numbers

`transform()` — [app.py:957](../app.py#L957)

| Field | How it's produced |
|---|---|
| `Currency` | Looked up from the company label in `COMPANY_CURRENCY` |
| `Reporting Month` | From the transaction date, formatted `6'2026` |
| `CostCenter` / `SubClass Name` | The QuickBooks class split on the first `:` |
| `Transaction Value in Original Currency` | `-Debit` if there's a debit, else `Credit`, else `Balance` |

**Debits are negative.** Revenue ends up positive and costs negative. Anything you add later has to
follow this, and it's why the AI assistant is told to use the absolute value of costs.

**Mapping** (`apply_mapping()`, [app.py:1027](../app.py#L1027)) pulls the leading number out of the
account name — `"110205 WISE RON"` → `110205` — and looks it up in the mapping sheet to add
`Mapping`, `Item` and `Statement`. `Item` is `P&L` or `B.S`, and everything downstream depends on it.

**Exchange rates** (`parse_forex_rate_file()`, [app.py:1102](../app.py#L1102)) are looked up by
currency and month:

- USD → always 1.0
- `P&L` rows → the **average** rate for that month
- balance sheet rows → the **closing** rate
- **nothing found → 1.0, with no warning**

That last line is the one that quietly ruins reports. Always check the rate file covers your dates.

---

## Pivot P&L

`build_pivot_report()` — [app.py:1541](../app.py#L1541). Uses only `P&L` rows.

**Rows** are a fixed hierarchy with subtotals worked in:

```
Revenue
HR Cost-COPS      → Total HR Cost-COPS
Other COPS
                  → Gross Profit   (Revenue − HR COPS − COPS)
                  → Gross Profit %
HR Cost-G&A       → Total HR Cost-G&A
Other Expenses
                  → Total Expenses
Other
```

Which bucket a statement lands in is **worked out, not configured**
(`_classify_statements()`, [app.py:1189](../app.py#L1189)): by name first (`HR Cost-COPS`,
`HR Cost-G&A`), otherwise by the most common leading digit of its account numbers — 4 is revenue,
5 is cost of sales, 6 is expenses, anything else is "other". Renaming statements or renumbering
accounts in the mapping sheet will move rows between sections, including into "other" where they
stop counting towards Gross Profit and Total Expenses.

**Columns** repeat for each section: three months (newest first) plus a variance column. Sections
are the consolidated total, then one per company, then one per cost center — with an option to break
`OPS50` down into its sub-classes.

`pivot_to_excel_bytes()` ([app.py:1642](../app.py#L1642)) writes the styled Excel version from the
raw numbers. The on-screen preview converts everything to text first, because the table mixes
numbers and percentages.

---

## Dashboard and AI assistant

`_render_financial_dashboard()` ([app.py:2592](../app.py#L2592)) — KPI tiles and six Plotly charts,
P&L only, all in USD.

The AI assistant uses Google Gemini and only runs if `GOOGLE_API_KEY` is set.
`_build_financial_context()` ([app.py:2373](../app.py#L2373)) summarises the data into text — totals
by company, month, mapping and cost center — and `_detect_and_filter()`
([app.py:2273](../app.py#L2273)) notices when you name a company and adds its detailed rows. Gemini
is instructed to answer only from that data. Both are cleared whenever a new report is generated, so
old numbers can't leak into new answers.

---

## Settings

Everything comes from `.streamlit/secrets.toml` — there are no environment variables.
`get_secret()` ([app.py:552](../app.py#L552)) and `_load_companies()`
([app.py:563](../app.py#L563)) read it, and both return empty rather than crashing if it's missing,
so a bad install shows up as "no data collected" plus a log line instead of a stack trace.
