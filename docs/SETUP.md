# Setup Guide

From nothing to a working report. About 15 minutes if you already have the credentials.

---

## 1. Install

You need **Python 3.11, 3.12 or 3.13**. Older versions will not work.

```bash
git clone https://github.com/hassanmohamed/lxt-finance-tool.git
cd lxt-finance-tool

python3.11 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## 2. Add your credentials

```bash
mkdir -p .streamlit
cp secrets.toml.example .streamlit/secrets.toml
```

Open `.streamlit/secrets.toml` and fill it in:

| Setting | Needed? | What it is |
|---|---|---|
| `APP_USERNAME` | Yes | The login email. |
| `APP_PASSWORD_HASH` | Yes | A bcrypt hash of the login password — not the password itself. Must start with `$2b$`. |
| `QB_CLIENT_ID` / `QB_CLIENT_SECRET` | Yes | From your Intuit app. One app covers all 9 companies. |
| `GOOGLE_API_KEY` | No | Turns on the AI assistant. Leave empty and everything else still works. |
| `[companies.*] label` | Yes | Company name. **Must match the code exactly** — see the table in the [README](../README.md#the-9-companies). |
| `[companies.*] realm_id` | Yes | The QuickBooks company ID. Never changes. |
| `[companies.*] refresh_token` | Yes | The QuickBooks token. **The app rewrites this after every run.** |

To set a new login password:

```bash
python -c "import bcrypt; print(bcrypt.hashpw(b'YOUR_PASSWORD', bcrypt.gensalt()).decode())"
```

> **Keep `.streamlit/secrets.toml` safe.** It holds live credentials for all 9 companies, it is
> gitignored, and it is the only copy of your QuickBooks tokens. Share it only through an encrypted
> channel, and back it up.

## 3. Run it

```bash
streamlit run app.py
```

Open <http://localhost:8501> and log in. Then follow
[Making a report](../README.md#making-a-report).

---

## Hosting it for a team

Run it on a server you control:

```bash
streamlit run app.py --server.port 8501 --server.address 0.0.0.0 --server.headless true
```

Put it behind a reverse proxy with **HTTPS** (nginx or Caddy) — the login sends a password and the
reports contain your full ledger. Keep it internal or behind a VPN if you can. Use `systemd` or
similar to keep it running.

Two rules that are not optional:

1. **`.streamlit/secrets.toml` must be writable and must survive restarts.** A container without a
   volume will lose your tokens.
2. **Only one instance may run against a given set of credentials.**

---

## Where credentials come from

| Credential | Source |
|---|---|
| `QB_CLIENT_ID` / `QB_CLIENT_SECRET` | [developer.intuit.com](https://developer.intuit.com) → My Apps → Keys & credentials (Production) |
| `realm_id` × 9 | The Intuit dashboard, or the existing secrets file. These never change. |
| `refresh_token` × 9 | The current `secrets.toml`. Making a new one means re-running Intuit's OAuth consent flow for that company. |
| `GOOGLE_API_KEY` | [aistudio.google.com/apikey](https://aistudio.google.com/apikey) |
| `APP_USERNAME` / `APP_PASSWORD_HASH` | Your choice — generate the hash with the command above. |

---

## Taking over the project

- [ ] Get `.streamlit/secrets.toml` over an encrypted channel, and confirm **the previous owner has
      stopped running their copy**.
- [ ] Install and generate one report end to end.
- [ ] Check all 9 companies succeeded — no errors listed, no "Token save failed" warning.
- [ ] Run it twice and confirm the `refresh_token` values in `secrets.toml` changed between runs.
      If they didn't, stop and fix that first.
- [ ] Compare one month's total USD revenue against the last report the previous team produced.
- [ ] Shut down every other copy of the app.
- [ ] Get an exchange-rate file that covers the current year.
- [ ] Set your own password.
- [ ] Get added to the Intuit developer app.
- [ ] Start backing up `.streamlit/secrets.toml`.
- [ ] Ask the previous owner to delete the old `lxt_qb_tokens.json` GitHub gist and revoke its
      access token. It used to hold these tokens and was readable by anyone with the link.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| "Password hash configuration error" | `APP_PASSWORD_HASH` isn't a real bcrypt hash. Regenerate it. |
| "Too many failed attempts" | 5 wrong logins locks you out for 15 minutes. Restarting the app clears it. |
| Logged out while working | Normal — 30 minutes of inactivity ends the session. |
| **"Token save failed"** | The app couldn't save the new token. **Fix immediately** — every run after this burns a token it can't save. Check that `.streamlit/secrets.toml` exists and is writable. |
| `invalid_grant` for one company | That token is dead, usually because another copy of the app used it. Re-run Intuit's OAuth consent flow for that company and paste the new token into `secrets.toml`. |
| All 9 companies fail | Wrong `QB_CLIENT_ID` / `QB_CLIENT_SECRET`, or sandbox credentials against live QuickBooks. |
| "No data was collected" | Every company failed — open the error list above the message. |
| Pivot P&L is empty | The mapping sheet wasn't uploaded, or has no rows marked `P&L`. |
| "Only N month(s) of data" | Your date range is under 3 months. Widen it. |
| A company's rate shows 1.0 | The rate file is missing that currency or month, or the company `label` doesn't match the code. |
| AI assistant unavailable | `GOOGLE_API_KEY` is empty. |
| Theme looks wrong | `.streamlit/config.toml` is missing from your clone. |
| Port already in use | `streamlit run app.py --server.port 8600` |
