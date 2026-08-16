# LXT Financial Consolidated Report

A password-protected Streamlit app that pulls the **General Ledger from 9 QuickBooks companies**,
converts everything to USD, and gives you:

- a consolidated **General Ledger Excel export**
- a **Pivot P&L** — 3 months side by side, with variance, per entity and per cost center
- a **dashboard** of charts
- an **AI assistant** you can ask questions about the data

Everything is in one file, [app.py](app.py). All settings are in `.streamlit/secrets.toml`.

**Docs:** [Setup guide](docs/SETUP.md) · [How it works](docs/HOW-IT-WORKS.md)

---

## Read this first

Three things cause almost every problem with this app.

**1. QuickBooks tokens are single-use.** Each run consumes the refresh tokens and writes new ones
back into `.streamlit/secrets.toml`. That file is the *only* copy.

> Only **one** copy of the app may ever run against these credentials, and it must be able to write
> to `secrets.toml`. Two copies will break each other.

**2. The exchange-rate file must cover your dates.** If a month is missing, the app uses a rate of
**1.0 and says nothing** — the report looks fine and the numbers are wrong.

> ⚠️ **Right now `ConsolidatedExchRate Accounting.csv` only goes up to January 2026.** Get an updated
> file from Finance before running any report covering later months.

**3. Company names in `secrets.toml` must match the code.** Each `label` has to match a key in
`COMPANY_CURRENCY` ([app.py:576](app.py#L576)) exactly, or that company gets no currency and a rate
of 1.0 — again, silently.

---

## Quick start

```bash
git clone https://github.com/hassanmohamed/lxt-finance-tool.git
cd lxt-finance-tool

python3.11 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt

mkdir -p .streamlit
cp secrets.toml.example .streamlit/secrets.toml
# fill in the credentials — see docs/SETUP.md

streamlit run app.py               # → http://localhost:8501
```

---

## Making a report

1. **Log in.**
2. **Pick your dates.** The Pivot P&L needs at least 3 months.
3. **Upload `Consol Mapping sheet.csv`** (in this repo) — tells the app which account is revenue,
   cost, or balance sheet.
4. **Upload `ConsolidatedExchRate Accounting.csv`** (in this repo) — the USD conversion rates.
   *Check it covers your months.*
5. **Click Generate.** It works through the 9 companies; any that fail are listed at the end.
6. **Download** the Excel, then scroll down for the Pivot P&L, charts, and AI chat.

---

## The 9 companies

| Key in secrets | Label | Currency |
|---|---|---|
| `lxt_egypt` | LXT Egypt | EGP |
| `lxt_canada` | LXT Canada | CAD |
| `lxt_australia` | LXT Australia | AUD |
| `lxt_romania` | LXT Romania | RON |
| `lxt_india` | LXT India | INR |
| `lxt_germany` | CW GmbH | EUR |
| `lxt_uk` | LXT UK | GBP |
| `lxt_usa` | LXT USA | USD |
| `lxt_clickworker_usa` | CW Inc | USD |

---

## Files

```
app.py                              The whole application
requirements.txt                    Dependencies
secrets.toml.example                Copy to .streamlit/secrets.toml and fill in
.streamlit/config.toml              Dark theme
.streamlit/secrets.toml             Your real credentials — never commit this
Consol Mapping sheet.csv            Account mapping (upload in the UI)
ConsolidatedExchRate Accounting.csv Exchange rates (upload in the UI)
assets/, docs/                      Logo, documentation
```

There is no test suite.
