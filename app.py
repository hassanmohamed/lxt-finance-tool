"""
LXT Financial Consolidated Report — Streamlit App
===================================================
Password-protected web app that extracts General Ledger data
from 9 QuickBooks Online companies and produces a downloadable
consolidated Excel report.
"""

import base64
import io
import os
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

# ─────────────────────────────────────────────────────────────
# Page Config
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="LXT Financial Consolidated Report",
    page_icon="📊",
    layout="wide",
)

# ─────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────
QB_TOKEN_URL = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"
QB_BASE_URL = "https://quickbooks.api.intuit.com"

# Path to the Consol Mapping sheet (lives alongside app.py)
MAPPING_CSV_PATH = Path(__file__).parent / "Consol Mapping sheet.csv"

# Path to the Streamlit secrets file (for auto-saving refresh tokens)
SECRETS_PATH = Path(__file__).parent / ".streamlit" / "secrets.toml"

# Company label → local currency (ISO codes)
COMPANY_CURRENCY = {
    "LXT Egypt": "EGP",
    "LXT Canada": "CAD",
    "LXT Australia": "AUD",
    "LXT Romania": "RON",
    "LXT India": "INR",
    "CW GmbH": "EUR",
    "LXT UK": "GBP",
    "LXT USA": "USD",
    "CW Inc": "USD",
}

# Currencies that need forex rate input (USD is always 1.0)
FOREX_CURRENCIES = ["EGP", "CAD", "AUD", "RON", "INR", "EUR", "GBP"]

OUTPUT_COLUMNS = [
    "Distribution account",
    "Account Number",
    "Transaction date",
    "Reporting Month",
    "Memo/Description",
    "Name",
    "Transaction id",
    "Customer full name",
    "Supplier",
    "Number",
    "Balance",
    "Debit",
    "Credit",
    "Class full name",
    "CostCenter",
    "SubClass Name",
    "Company Country",
    "Mapping",
    "Item",
    "Statement",
    "Transaction Value in Original Currency",
    "Currency",
    "Forex Rate",
    "Amount in USD (Reporting Currency)",
]

QB_COLUMN_MAP = {
    # Internal API keys
    "account_name": "Distribution account",
    "tx_date": "Transaction date",
    "memo": "Memo/Description",
    "name": "Name",
    "txn_type": "Transaction id",
    "cust_name": "Customer full name",
    "vend_name": "Supplier",
    "doc_num": "Number",
    "subt_nat_amount": "Balance",
    "subt_nat_home_amount": "Balance",
    "debt_amt": "Debit",
    "credit_amt": "Credit",
    "klass_name": "Class full name",

    # Display-name ColTitles returned by the API
    "Account": "Distribution account",
    "Distribution Account": "Distribution account",
    "Transaction Type": "Transaction id",
    "Trans #": "Transaction id",
    "No.": "Number",
    "Num": "Number",
    "Customer": "Customer full name",
    "Vendor": "Supplier",
    "Memo/Description": "Memo/Description",
    "Date": "Transaction date",
    "Class": "Class full name",

    "Amount": "Balance",
    "Debit": "Debit",
    "Credit": "Credit",
    # Home-currency debit/credit (from debt_home_amt / credit_home_amt)
    "Home Debit": "Debit",
    "Home Credit": "Credit",
    "debt_home_amt": "Debit",
    "credit_home_amt": "Credit",
    # For multi-currency companies the API labels transaction-currency
    # amounts as "Foreign Debit"/"Foreign Credit". We intentionally do NOT
    # map them to Debit/Credit — we want home-currency amounts instead.
    # "Foreign Debit" and "Foreign Credit" are kept as-is and ignored.
}

QB_REPORT_COLUMNS = "account_name,tx_date,memo,name,txn_type,cust_name,vend_name,doc_num,subt_nat_amount,subt_nat_home_amount,debt_amt,credit_amt,debt_home_amt,credit_home_amt,klass_name"


# ═══════════════════════════════════════════════════════════════
# Authentication UI
# ═══════════════════════════════════════════════════════════════
def check_password() -> bool:
    """Show a login form and return True if authenticated."""
    if st.session_state.get("authenticated"):
        return True

    st.markdown(
        """
        <div style="display:flex; justify-content:center; margin-top:8vh;">
            <div style="
                background: linear-gradient(135deg, #1e1e2f 0%, #2d2d44 100%);
                border: 1px solid #3a3a5c;
                border-radius: 16px;
                padding: 3rem 2.5rem;
                width: 400px;
                box-shadow: 0 8px 32px rgba(0,0,0,0.4);
            ">
                <h2 style="text-align:center; margin-bottom:0.2rem; color:#fff;">
                    🔐 LXT Reports
                </h2>
                <p style="text-align:center; color:#888; font-size:0.9rem; margin-bottom:2rem;">
                    Enter your password to continue
                </p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        password = st.text_input(
            "Password",
            type="password",
            placeholder="Enter password…",
            label_visibility="collapsed",
        )
        login_clicked = st.button("Login", width="stretch", type="primary")

        if login_clicked:
            if password == st.secrets["APP_PASSWORD"]:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("❌ Incorrect password. Please try again.")

    return False


# ═══════════════════════════════════════════════════════════════
# QuickBooks API Functions
# ═══════════════════════════════════════════════════════════════
def refresh_access_token(
    client_id: str, client_secret: str, refresh_token: str
) -> dict:
    """Exchange a refresh token for a fresh access token."""
    credentials = f"{client_id}:{client_secret}"
    auth_header = base64.b64encode(credentials.encode()).decode()

    resp = requests.post(
        QB_TOKEN_URL,
        headers={
            "Accept": "application/json",
            "Authorization": f"Basic {auth_header}",
            "Content-Type": "application/x-www-form-urlencoded",
        },
        data={"grant_type": "refresh_token", "refresh_token": refresh_token},
        timeout=30,
    )

    if resp.status_code != 200:
        raise RuntimeError(
            f"Token refresh failed (HTTP {resp.status_code}): {resp.text}"
        )

    data = resp.json()
    return {
        "access_token": data["access_token"],
        "refresh_token": data.get("refresh_token", refresh_token),
    }

# ═══════════════════════════════════════════════════════════════
# GitHub Gist — Persistent Token Storage
# ═══════════════════════════════════════════════════════════════
GIST_FILENAME = "lxt_qb_tokens.json"
GIST_API = "https://api.github.com/gists"


def _get_github_headers() -> dict:
    """Return GitHub API headers using the token from secrets."""
    token = st.secrets.get("GITHUB_TOKEN", "")
    return {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
    }


def _find_token_gist() -> str | None:
    """Find the existing private gist ID by filename, or return None."""
    try:
        resp = requests.get(GIST_API, headers=_get_github_headers(), timeout=15)
        if resp.status_code == 200:
            for gist in resp.json():
                if GIST_FILENAME in gist.get("files", {}):
                    return gist["id"]
    except Exception:
        pass
    return None


def _load_gist_tokens() -> dict | None:
    """Load refresh tokens from the private gist. Returns dict {company_key: token}."""
    gist_id = _find_token_gist()
    if not gist_id:
        return None
    try:
        resp = requests.get(
            f"{GIST_API}/{gist_id}", headers=_get_github_headers(), timeout=15
        )
        if resp.status_code == 200:
            import json
            content = resp.json()["files"][GIST_FILENAME]["content"]
            return json.loads(content)
    except Exception:
        pass
    return None


def _save_gist_tokens(tokens: dict) -> None:
    """Create or update the private gist with current tokens."""
    import json
    payload = {
        "description": "LXT QuickBooks refresh tokens (auto-managed)",
        "files": {
            GIST_FILENAME: {"content": json.dumps(tokens, indent=2)}
        },
    }
    try:
        gist_id = _find_token_gist()
        headers = _get_github_headers()
        if gist_id:
            requests.patch(
                f"{GIST_API}/{gist_id}",
                headers=headers,
                json=payload,
                timeout=15,
            )
        else:
            payload["public"] = False
            requests.post(
                GIST_API,
                headers=headers,
                json=payload,
                timeout=15,
            )
    except Exception:
        pass


def _save_refresh_token(old_token: str, new_token: str) -> None:
    """
    Replace the old refresh token with the new one in secrets.toml (local).
    This ensures the next run uses the latest single-use token.
    """
    if old_token == new_token or not SECRETS_PATH.exists():
        return

    try:
        content = SECRETS_PATH.read_text()
        if old_token in content:
            content = content.replace(old_token, new_token)
            SECRETS_PATH.write_text(content)
    except Exception:
        pass


def fetch_general_ledger(
    access_token: str, realm_id: str, start_date: str, end_date: str
) -> dict:
    """GET the General Ledger report JSON from QBO."""
    url = f"{QB_BASE_URL}/v3/company/{realm_id}/reports/GeneralLedger"
    resp = requests.get(
        url,
        headers={
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
        },
        params={
            "start_date": start_date,
            "end_date": end_date,
            "columns": QB_REPORT_COLUMNS,
            "accounting_method": "Accrual",
        },
        timeout=120,
    )

    if resp.status_code != 200:
        raise RuntimeError(
            f"API request failed (HTTP {resp.status_code}): {resp.text}"
        )

    return resp.json()


# ═══════════════════════════════════════════════════════════════
# JSON Flattening (Recursive)
# ═══════════════════════════════════════════════════════════════
def flatten_report(report_json: dict) -> list[dict]:
    """Flatten the nested QBO report JSON into a list of row dicts."""
    columns_meta = report_json.get("Columns", {}).get("Column", [])
    col_keys = [c.get("ColTitle", "").strip() for c in columns_meta]

    rows = report_json.get("Rows", {}).get("Row", [])
    flat: list[dict] = []
    _walk(rows, col_keys, flat)
    return flat


def _walk(rows: list, col_keys: list[str], acc: list[dict]) -> None:
    """Recursively collect only type='Data' rows."""
    for row in rows:
        rtype = row.get("type", "Data")

        if rtype == "Data":
            cells = row.get("ColData", [])
            record = {
                col_keys[i] if i < len(col_keys) else f"col_{i}": c.get("value", "")
                for i, c in enumerate(cells)
            }
            acc.append(record)

        elif rtype == "Section":
            nested = row.get("Rows", {}).get("Row", [])
            if nested:
                _walk(nested, col_keys, acc)


# ═══════════════════════════════════════════════════════════════
# Pandas Transformation
# ═══════════════════════════════════════════════════════════════
def transform(raw_rows: list[dict], company_label: str) -> pd.DataFrame:
    """Rename columns, add Company Country, filter nulls."""
    if not raw_rows:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    df = pd.DataFrame(raw_rows)

    # Rename columns using the mapping.
    # For multi-currency companies the API returns "Foreign Debit" / "Foreign Credit"
    # and "Amount" (from subt_nat_home_amount) — all in the company's home currency.
    # The QB_COLUMN_MAP handles all variants (Foreign/Nat/plain) → Debit/Credit/Balance.
    rename_map = {
        k: v for k, v in QB_COLUMN_MAP.items() if k in df.columns and k != v
    }
    df = df.rename(columns=rename_map)

    # Enrich Transaction id
    if "Transaction id" in df.columns and "Number" in df.columns:
        df["Transaction id"] = (
            df["Transaction id"].astype(str).str.strip()
            + " #"
            + df["Number"].astype(str).str.strip()
        )

    # Ensure all columns exist
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df["Company Country"] = company_label
    df["Currency"] = COMPANY_CURRENCY.get(company_label, "")

    # Split "Class full name" on ":" into CostCenter and SubClass Name
    class_split = df["Class full name"].astype(str).str.split(":", n=1, expand=True)
    df["CostCenter"] = class_split[0].str.strip() if 0 in class_split.columns else ""
    df["SubClass Name"] = class_split[1].str.strip() if 1 in class_split.columns else ""

    # Reporting Month: month'year (e.g. "2'2026") derived from Transaction date
    td = pd.to_datetime(df["Transaction date"], format="mixed", errors="coerce")
    df["Reporting Month"] = td.dt.month.astype("Int64").astype(str) + "'" + td.dt.year.astype("Int64").astype(str)

    df = df[OUTPUT_COLUMNS]

    # Numeric conversion
    for c in ("Balance", "Debit", "Credit"):
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # Transaction Value in Original Currency:
    # If Debit has a value → -Debit, else if Credit has a value → Credit, else Balance
    import numpy as np
    df["Transaction Value in Original Currency"] = np.where(
        df["Debit"].notna(), -df["Debit"],
        np.where(df["Credit"].notna(), df["Credit"], df["Balance"])
    )

    # Filter empty Distribution account
    df["Distribution account"] = df["Distribution account"].astype(str).str.strip()
    df = df[
        (df["Distribution account"] != "")
        & (df["Distribution account"].str.lower() != "none")
        & (df["Distribution account"].str.lower() != "nan")
    ]
    df = df.dropna(subset=["Distribution account"]).reset_index(drop=True)

    return df


# ═══════════════════════════════════════════════════════════════
# Consol Mapping Sheet Lookup
# ═══════════════════════════════════════════════════════════════
@st.cache_data
def load_mapping() -> pd.DataFrame:
    """Load the Consol Mapping sheet and return a lookup DataFrame."""
    if not MAPPING_CSV_PATH.exists():
        st.warning(f"Mapping file not found: {MAPPING_CSV_PATH}")
        return pd.DataFrame(columns=["Account Number", "Mapping", "Item", "Statement"])

    mapping_df = pd.read_csv(MAPPING_CSV_PATH)
    # Normalise column names (strip whitespace)
    mapping_df.columns = mapping_df.columns.str.strip()
    # Ensure Account Number is string without float decimals (pandas reads as float64)
    mapping_df["Account Number"] = (
        mapping_df["Account Number"]
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.strip()
    )
    # Keep only the columns we need; drop duplicates so first match wins
    mapping_df = mapping_df[["Account Number", "Mapping", "Item", "Statement"]]
    mapping_df = mapping_df.drop_duplicates(subset="Account Number", keep="first")
    return mapping_df


def apply_mapping(df: pd.DataFrame) -> pd.DataFrame:
    """
    Extract the leading account number code from 'Distribution account'
    (e.g. '110205' from '110205 WISE RON') and merge with the Consol
    Mapping sheet to add Mapping, Item, and Statement columns.
    """
    mapping_df = load_mapping()
    if mapping_df.empty:
        return df

    # Extract leading numeric code from Distribution account
    df["_account_code"] = (
        df["Distribution account"]
        .astype(str)
        .str.extract(r"^(\d+)", expand=False)
        .str.strip()
    )

    # Merge on the account code
    df = df.merge(
        mapping_df,
        left_on="_account_code",
        right_on="Account Number",
        how="left",
        suffixes=("", "_map"),
    )

    # Keep Account Number from the mapping (the code used for lookup)
    if "Account Number_map" in df.columns:
        df["Account Number"] = df["Account Number_map"]
        df = df.drop(columns=["Account Number_map"])
    elif "Account Number" not in df.columns:
        df["Account Number"] = df["_account_code"]

    # If Mapping/Item/Statement columns already existed (as empty placeholders),
    # overwrite them with the merged values
    for col in ("Mapping", "Item", "Statement"):
        if f"{col}_map" in df.columns:
            df[col] = df[f"{col}_map"]
            df = df.drop(columns=[f"{col}_map"])

    # Clean up helper columns
    df = df.drop(columns=["_account_code"], errors="ignore")

    return df


# ═══════════════════════════════════════════════════════════════
# Forex Rate File Parser
# ═══════════════════════════════════════════════════════════════
def parse_forex_rate_file(uploaded_file) -> dict[str, dict[str, dict[str, float]]]:
    """
    Parse an uploaded Exchange Rate file (Excel or CSV) and return a
    nested lookup:
        { currency: { month_key: { "closing": rate, "average": rate } } }

    Expected columns:
      A  Currency
      B  End of Month       (date — used to derive month'year key)
      F  ClosingRate2       (closing rate vs USD)
      G  AverageRate2       (average rate vs USD)
    """
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    # Normalise column names
    df.columns = df.columns.str.strip()

    # Validate required columns exist
    required = {"Currency", "End of Month", "ClosingRate2", "AverageRate2"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Uploaded file is missing columns: {', '.join(missing)}")

    # Parse the End of Month date and build a month key like "1'2026"
    df["_eom"] = pd.to_datetime(df["End of Month"], format="mixed", errors="coerce")
    df["_month_key"] = (
        df["_eom"].dt.month.astype("Int64").astype(str)
        + "'"
        + df["_eom"].dt.year.astype("Int64").astype(str)
    )

    df["ClosingRate2"] = pd.to_numeric(df["ClosingRate2"], errors="coerce").fillna(1.0)
    df["AverageRate2"] = pd.to_numeric(df["AverageRate2"], errors="coerce").fillna(1.0)

    # Build lookup dict
    forex: dict[str, dict[str, dict[str, float]]] = {}
    for _, row in df.iterrows():
        ccy = str(row["Currency"]).strip().upper()
        mk = str(row["_month_key"]).strip()
        if not ccy or not mk or mk == "<NA>'<NA>":
            continue
        forex.setdefault(ccy, {})[mk] = {
            "closing": float(row["ClosingRate2"]),
            "average": float(row["AverageRate2"]),
        }

    # Ensure USD always resolves to 1.0
    if "USD" not in forex:
        forex["USD"] = {}

    return forex


# ═══════════════════════════════════════════════════════════════
# Excel Export (in-memory)
# ═══════════════════════════════════════════════════════════════
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Write DataFrame to an in-memory Excel file and return bytes."""
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════
# Pivot P&L Report
# ═══════════════════════════════════════════════════════════════
def _month_key(year: int, month: int) -> str:
    """Return the Reporting Month key matching the format in the data, e.g. '1\'2026'."""
    return f"{month}'{year}"


def _prev_months(year: int, month: int, count: int) -> list[tuple[int, int]]:
    """Return a list of (year, month) tuples going back `count` months inclusive."""
    result = []
    for _ in range(count):
        result.append((year, month))
        month -= 1
        if month == 0:
            month = 12
            year -= 1
    return result


def _build_row_index(master_df: pd.DataFrame) -> list[dict]:
    """
    Build the shared row hierarchy: Statement (group) → Mapping (detail) + totals.
    Each entry has: Code, Description, _style, _statement, _mapping.
    """
    rows: list[dict] = []
    statements = master_df["Statement"].dropna().unique()

    for stmt in statements:
        stmt_str = str(stmt).strip()
        if not stmt_str or stmt_str.lower() == "nan":
            continue

        stmt_df = master_df[master_df["Statement"] == stmt]

        # Statement group header
        rows.append({
            "Code": "",
            "Description": stmt_str,
            "_style": "group",
            "_statement": stmt,
            "_mapping": None,
        })

        # Mapping detail lines
        mappings = stmt_df["Mapping"].dropna().unique()
        for mapping in mappings:
            mapping_str = str(mapping).strip()
            if not mapping_str or mapping_str.lower() == "nan":
                continue

            # Find the most common Account Number for this mapping
            map_df = stmt_df[stmt_df["Mapping"] == mapping]
            acct_nums = map_df["Account Number"].dropna().astype(str).str.strip()
            acct_nums = acct_nums[acct_nums != ""]
            code = acct_nums.mode().iloc[0] if len(acct_nums) > 0 else ""

            rows.append({
                "Code": code,
                "Description": f"  {mapping_str}",
                "_style": "detail",
                "_statement": stmt,
                "_mapping": mapping,
            })

        # Statement total row
        rows.append({
            "Code": "",
            "Description": f"Total {stmt_str}",
            "_style": "total",
            "_statement": stmt,
            "_mapping": "__TOTAL__",
        })

    return rows


def _compute_section_values(
    df: pd.DataFrame,
    row_index: list[dict],
    month_keys: list[str],
    col_prefix: str,
    month_labels: list[str],
) -> tuple[list[str], dict[int, dict[str, float]]]:
    """
    Compute values for one section (Consolidated / Entity / CostCenter).

    Returns:
      - col_names: list of column names [prefix M1, prefix M2, prefix M3, prefix Var]
      - values: dict mapping row_index position → {col_name: value}
    """
    col_names = [f"{col_prefix} {lbl}" for lbl in month_labels] + [f"{col_prefix} Variance"]
    values: dict[int, dict[str, float]] = {}

    # Track totals per statement for total rows
    stmt_totals: dict[str, dict[str, float]] = {}

    for idx, row in enumerate(row_index):
        style = row["_style"]
        stmt = row["_statement"]
        mapping = row["_mapping"]

        if style == "group":
            # Group header — no values (use None for numeric compatibility)
            values[idx] = {c: None for c in col_names}
            continue

        if style == "detail" and mapping is not None:
            row_vals = {}
            for i, mk in enumerate(month_keys):
                cn = col_names[i]
                mask = (
                    (df["Statement"] == stmt)
                    & (df["Mapping"] == mapping)
                    & (df["Reporting Month"] == mk)
                )
                val = df.loc[mask, "Amount in USD (Reporting Currency)"].sum()
                row_vals[cn] = round(val, 2)

                # Accumulate statement totals
                total_key = str(stmt)
                if total_key not in stmt_totals:
                    stmt_totals[total_key] = {c: 0.0 for c in col_names[:-1]}
                stmt_totals[total_key][cn] += val

            # Variance = latest month - previous month
            row_vals[col_names[-1]] = round(row_vals[col_names[0]] - row_vals[col_names[1]], 2)
            values[idx] = row_vals

        elif style == "total" and mapping == "__TOTAL__":
            total_key = str(stmt)
            totals = stmt_totals.get(total_key, {c: 0.0 for c in col_names[:-1]})
            row_vals = {c: round(totals.get(c, 0.0), 2) for c in col_names[:-1]}
            row_vals[col_names[-1]] = round(row_vals[col_names[0]] - row_vals[col_names[1]], 2)
            values[idx] = row_vals

    return col_names, values


def build_pivot_report(
    master_df: pd.DataFrame,
    selected_year: int,
    selected_month: int,
) -> tuple[pd.DataFrame, list[dict], list[str], list[tuple[str, list[str]]]]:
    """
    Build the full horizontal pivot P&L report.

    Returns:
      - display_df: DataFrame for Streamlit preview
      - raw_rows: list of row dicts with '_style' metadata
      - all_columns: ordered list of all column names
      - section_groups: list of (section_label, [col_names]) for Excel header grouping
    """
    import calendar

    # Determine 3 consecutive months (latest first)
    months = _prev_months(selected_year, selected_month, 3)
    month_keys = [_month_key(y, m) for y, m in months]
    month_labels = [calendar.month_abbr[m] + f" {y}" for y, m in months]

    pl_df = master_df.copy()

    # Build shared row hierarchy
    row_index = _build_row_index(pl_df)

    # Collect all section column groups
    all_section_cols: list[str] = []
    section_groups: list[tuple[str, list[str]]] = []

    # Helper to add a section
    def add_section(data_df, prefix):
        col_names, vals = _compute_section_values(
            data_df, row_index, month_keys, prefix, month_labels
        )
        all_section_cols.extend(col_names)
        section_groups.append((prefix, col_names))
        return vals

    # ── Section 1: Consolidated ──
    all_values = [add_section(pl_df, "Consolidated")]

    # ── Section 2: Per Legal Entity ──
    entities = sorted(pl_df["Company Country"].dropna().unique())
    for entity in entities:
        entity_df = pl_df[pl_df["Company Country"] == entity]
        all_values.append(add_section(entity_df, entity))

    # ── Section 3: Per CostCenter ──
    cost_centers = pl_df["CostCenter"].dropna().astype(str).str.strip()
    cost_centers = sorted(cost_centers[cost_centers != ""].unique())
    for cc in cost_centers:
        cc_df = pl_df[pl_df["CostCenter"].astype(str).str.strip() == cc]
        all_values.append(add_section(cc_df, f"CC: {cc}"))

    # Build final rows
    all_columns = ["Code", "Description"] + all_section_cols
    raw_rows: list[dict] = []

    for idx, row in enumerate(row_index):
        out = {
            "Code": row["Code"],
            "Description": row["Description"],
            "_style": row["_style"],
        }
        for section_vals in all_values:
            if idx in section_vals:
                out.update(section_vals[idx])
            # Missing values default to empty for group rows
        raw_rows.append(out)

    # Build display DataFrame
    display_df = pd.DataFrame(raw_rows)
    display_cols = [c for c in all_columns if c in display_df.columns]
    display_df = display_df[display_cols]

    return display_df, raw_rows, all_columns, section_groups


def pivot_to_excel_bytes(
    rows: list[dict],
    columns: list[str],
    section_groups: list[tuple[str, list[str]]],
    month_labels: list[str],
) -> bytes:
    """Write the horizontal pivot report to a styled Excel file."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Pivot P&L Report"

    # ── Styles ──
    section_font = Font(bold=True, size=11, color="FFFFFF")
    section_fill = PatternFill(start_color="2D2D44", end_color="2D2D44", fill_type="solid")
    group_font = Font(bold=True, size=10, color="1B3A5C")
    group_fill = PatternFill(start_color="E8EEF4", end_color="E8EEF4", fill_type="solid")
    total_font = Font(bold=True, size=10)
    total_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    header_font = Font(bold=True, size=10, color="FFFFFF")
    header_fill = PatternFill(start_color="3A3A5C", end_color="3A3A5C", fill_type="solid")
    section_header_fills = [
        PatternFill(start_color="1B4F72", end_color="1B4F72", fill_type="solid"),
        PatternFill(start_color="7D3C98", end_color="7D3C98", fill_type="solid"),
        PatternFill(start_color="1E8449", end_color="1E8449", fill_type="solid"),
        PatternFill(start_color="B9770E", end_color="B9770E", fill_type="solid"),
        PatternFill(start_color="922B21", end_color="922B21", fill_type="solid"),
    ]
    num_fmt = '#,##0.00'
    center_align = Alignment(horizontal="center")
    right_align = Alignment(horizontal="right")
    thin_border = Border(
        bottom=Side(style="thin", color="CCCCCC"),
    )

    # ── Row 1: Section group headers (merged) ──
    # Code + Description stay empty on row 1
    ws.cell(row=1, column=1, value="Code").font = header_font
    ws.cell(row=1, column=1).fill = header_fill
    ws.cell(row=1, column=1).alignment = center_align
    ws.cell(row=1, column=2, value="Description").font = header_font
    ws.cell(row=1, column=2).fill = header_fill
    ws.cell(row=1, column=2).alignment = center_align

    col_offset = 3  # sections start at column C
    for sec_idx, (sec_label, sec_cols) in enumerate(section_groups):
        fill = section_header_fills[sec_idx % len(section_header_fills)]
        start_col = col_offset
        end_col = col_offset + len(sec_cols) - 1

        # Merge section header
        ws.merge_cells(
            start_row=1, start_column=start_col,
            end_row=1, end_column=end_col
        )
        cell = ws.cell(row=1, column=start_col, value=sec_label)
        cell.font = Font(bold=True, size=11, color="FFFFFF")
        cell.fill = fill
        cell.alignment = center_align

        col_offset += len(sec_cols)

    # ── Row 2: Month sub-headers ──
    ws.cell(row=2, column=1, value="").fill = header_fill
    ws.cell(row=2, column=2, value="").fill = header_fill

    col_offset = 3
    for sec_idx, (sec_label, sec_cols) in enumerate(section_groups):
        for i, col_name in enumerate(sec_cols):
            cell = ws.cell(row=2, column=col_offset + i)
            # Extract the sub-label (remove the section prefix)
            parts = col_name.split(" ", 1)
            if len(parts) > 1:
                # Remove the prefix before the month label
                prefix = sec_label + " "
                sub_label = col_name[len(prefix):] if col_name.startswith(prefix) else col_name
            else:
                sub_label = col_name
            cell.value = sub_label
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
        col_offset += len(sec_cols)

    # ── Data rows (starting row 3) ──
    for row_idx, row_data in enumerate(rows, start=3):
        style = row_data.get("_style", "detail")

        for col_idx, col_name in enumerate(columns, start=1):
            val = row_data.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)

            if style == "group":
                cell.font = group_font
                cell.fill = group_fill
            elif style == "total":
                cell.font = total_font
                cell.fill = total_fill
            elif style == "detail":
                cell.border = thin_border

            # Number formatting for value columns (col 3+)
            if col_idx > 2 and isinstance(val, (int, float)):
                cell.number_format = num_fmt
                cell.alignment = right_align

    # ── Column widths ──
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 38
    for col_idx in range(3, len(columns) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 15

    # Freeze panes: freeze Code + Description columns and the 2 header rows
    ws.freeze_panes = "C3"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════
# Main App
# ═══════════════════════════════════════════════════════════════
def main_app():
    """Render the main application after authentication."""

    # ── Sidebar ───────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### 👤 Admin")
        st.caption("Logged in as **Admin**")
        st.divider()
        if st.button("🚪 Logout", width="stretch"):
            st.session_state.clear()
            st.rerun()

    # ── Header ────────────────────────────────────────────────
    st.title("📊 LXT Financial Consolidated Report")
    st.markdown(
        "Extract General Ledger data from **9 QuickBooks companies** "
        "and download a single consolidated Excel report."
    )
    st.divider()

    # ── Date Inputs ───────────────────────────────────────────
    today = date.today()
    first_of_month = today.replace(day=1)

    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        start_date = st.date_input("📅 Start Date", value=first_of_month)
    with col2:
        end_date = st.date_input("📅 End Date", value=today)

    if start_date > end_date:
        st.error("Start date cannot be after end date.")
        return

    st.divider()

    # ── Forex Rate File Upload ────────────────────────────────
    with st.expander("💱 Exchange Rate File", expanded=True):
        st.caption(
            "Upload the **Consolidated Exchange Rate** file (Excel or CSV). "
            "The file should contain columns: **Currency** (A), "
            "**End of Month** (B), **ClosingRate2** (F), **AverageRate2** (G). "
            "Rates will be matched automatically by currency and month."
        )
        forex_file = st.file_uploader(
            "Upload Exchange Rate File",
            type=["xlsx", "xls", "csv"],
            key="forex_file_upload",
            label_visibility="collapsed",
        )

        forex_rates: dict[str, dict[str, dict[str, float]]] = {}
        if forex_file is not None:
            try:
                forex_rates = parse_forex_rate_file(forex_file)
                # Show summary of loaded rates
                currencies_loaded = sorted(forex_rates.keys())
                months_count = max(len(v) for v in forex_rates.values()) if forex_rates else 0
                st.success(
                    f"✅ Loaded rates for **{len(currencies_loaded)}** currencies "
                    f"across **{months_count}** months: {', '.join(currencies_loaded)}"
                )
            except Exception as exc:
                st.error(f"❌ Failed to parse exchange rate file: {exc}")
                forex_rates = {}
        else:
            st.info("ℹ️ No exchange rate file uploaded — all forex rates will default to **1.0**.")

    st.divider()

    # ── Generate Button ─────────────────────────────────────
    generate = st.button(
        "🚀 Generate Report",
        type="primary",
        width="stretch",
    )

    if generate:
        _run_etl(
            start_date.strftime("%Y-%m-%d"),
            end_date.strftime("%Y-%m-%d"),
            forex_rates,
        )

    # ── Show report results (persisted in session state) ──────
    if "report_data" in st.session_state and "report_name" in st.session_state:
        st.divider()

        col1, col2 = st.columns(2)
        col1.metric("Total Rows", f"{st.session_state['report_rows']:,}")
        col2.metric("File", st.session_state["report_name"])

        with st.expander("📋 Preview Data (first 100 rows)", expanded=True):
            st.dataframe(st.session_state["report_preview"], width="stretch")

        # Ensure filename is a plain string with .xlsx extension
        fname = str(st.session_state["report_name"])
        if not fname.endswith(".xlsx"):
            fname += ".xlsx"

        # Download button with Content-Disposition via Streamlit
        st.download_button(
            label="📥 Download Excel Report",
            data=st.session_state["report_data"],
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="download_report_btn",
        )

    # ── Pivot P&L Report Section ─────────────────────────────
    if "master_df" in st.session_state:
        st.divider()
        st.subheader("📈 Pivot P&L Report")
        st.markdown(
            "Generate a pivot report grouped by **Statement → Mapping**, "
            "across **3 consecutive months** with variance, split by "
            "**Legal Entity** and **CostCenter**."
        )

        p_col1, p_col2 = st.columns([1, 1])
        with p_col1:
            pivot_month = st.selectbox(
                "📅 Latest Month",
                options=list(range(1, 13)),
                format_func=lambda m: [
                    "January", "February", "March", "April", "May", "June",
                    "July", "August", "September", "October", "November", "December"
                ][m - 1],
                index=0,
                key="pivot_month",
            )
        with p_col2:
            pivot_year = st.number_input(
                "📅 Year",
                min_value=2020,
                max_value=2030,
                value=2026,
                step=1,
                key="pivot_year",
            )

        # Show which 3 months will be used
        import calendar
        months_preview = _prev_months(int(pivot_year), int(pivot_month), 3)
        months_str = ", ".join(
            f"{calendar.month_abbr[m]} {y}" for y, m in months_preview
        )
        st.caption(f"Report months: **{months_str}**")

        gen_pivot = st.button(
            "📈 Generate Pivot Report",
            type="primary",
            key="gen_pivot_btn",
        )

        if gen_pivot:
            with st.spinner("Building pivot report…"):
                display_df, raw_rows, all_columns, section_groups = build_pivot_report(
                    st.session_state["master_df"],
                    int(pivot_year),
                    int(pivot_month),
                )
                # Get month labels for Excel sub-headers
                import calendar as _cal
                _m_tuples = _prev_months(int(pivot_year), int(pivot_month), 3)
                _m_labels = [_cal.month_abbr[m] + f" {y}" for y, m in _m_tuples]

                pivot_xlsx = pivot_to_excel_bytes(
                    raw_rows, all_columns, section_groups, _m_labels
                )

                m_labels = [
                    _cal.month_abbr[m] + f"{y}"
                    for y, m in months_preview
                ]
                pivot_fname = f"LXT_Pivot_PL_{m_labels[0]}_to_{m_labels[2]}.xlsx"

                st.session_state["pivot_data"] = pivot_xlsx
                st.session_state["pivot_name"] = pivot_fname
                st.session_state["pivot_preview"] = display_df
                st.session_state["pivot_rows"] = len(display_df)

        # Show persisted pivot results
        if "pivot_data" in st.session_state:
            st.divider()

            col1, col2 = st.columns(2)
            col1.metric("Pivot Rows", f"{st.session_state['pivot_rows']:,}")
            col2.metric("File", st.session_state["pivot_name"])

            with st.expander("📋 Pivot Preview (first 200 rows)", expanded=True):
                st.dataframe(
                    st.session_state["pivot_preview"].head(200),
                    width="stretch",
                )

            pivot_fname = str(st.session_state["pivot_name"])
            if not pivot_fname.endswith(".xlsx"):
                pivot_fname += ".xlsx"

            st.download_button(
                label="📥 Download Pivot Report",
                data=st.session_state["pivot_data"],
                file_name=pivot_fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="download_pivot_btn",
            )


def _run_etl(start_date: str, end_date: str, forex_rates: dict):
    """Execute the full ETL pipeline with progress UI."""

    # Load credentials from secrets
    client_id = st.secrets["QB_CLIENT_ID"]
    client_secret = st.secrets["QB_CLIENT_SECRET"]
    companies = st.secrets["companies"]

    # Load latest tokens from GitHub Gist (falls back to secrets.toml)
    gist_tokens = _load_gist_tokens() or {}
    updated_tokens: dict[str, str] = {}

    all_frames: list[pd.DataFrame] = []
    errors: list[str] = []

    with st.spinner("Fetching data from QuickBooks…"):
        progress = st.progress(0, text="Starting…")
        status_container = st.status(
            "Processing 9 companies…", expanded=True
        )

        company_keys = list(companies.keys())
        total = len(company_keys)

        for idx, key in enumerate(company_keys):
            company = companies[key]
            label = company["label"]
            realm_id = company["realm_id"]
            # Prefer gist token, fall back to secrets.toml
            refresh_token = gist_tokens.get(key, company["refresh_token"])

            progress.progress(
                (idx) / total,
                text=f"Processing {label} ({idx + 1}/{total})…",
            )

            try:
                with status_container:
                    st.write(f"🔄 **{label}** — Authenticating…")

                # Auth — try gist token first, fall back to secrets.toml
                token_info = None
                secrets_token = company["refresh_token"]
                used_token = refresh_token
                try:
                    token_info = refresh_access_token(
                        client_id, client_secret, refresh_token
                    )
                except RuntimeError as auth_err:
                    # If gist token failed and we have a different secrets.toml token, retry
                    if (
                        "invalid_grant" in str(auth_err)
                        and secrets_token
                        and secrets_token != refresh_token
                    ):
                        with status_container:
                            st.write(
                                f"⚠️ **{label}** — Gist token expired, "
                                f"retrying with secrets.toml token…"
                            )
                        token_info = refresh_access_token(
                            client_id, client_secret, secrets_token
                        )
                        used_token = secrets_token
                    else:
                        raise

                # Track the new refresh token
                new_refresh = token_info["refresh_token"]
                updated_tokens[key] = new_refresh

                # Also save locally (best-effort)
                if new_refresh != used_token:
                    _save_refresh_token(used_token, new_refresh)

                with status_container:
                    st.write(f"📥 **{label}** — Fetching General Ledger…")

                # Extract
                report_json = fetch_general_ledger(
                    token_info["access_token"], realm_id, start_date, end_date
                )
                raw_rows = flatten_report(report_json)

                # Transform
                df = transform(raw_rows, label)

                with status_container:
                    st.write(f"✅ **{label}** — {len(df)} rows extracted.")

                if not df.empty:
                    all_frames.append(df)

            except Exception as exc:
                msg = f"❌ **{label}** — {exc}"
                errors.append(msg)
                with status_container:
                    st.write(msg)

        progress.progress(1.0, text="Done!")
        status_container.update(
            label="Processing complete!", state="complete", expanded=False
        )

    # ── Save updated tokens to GitHub Gist ─────────────────────
    if updated_tokens:
        _save_gist_tokens(updated_tokens)

    # ── Results ───────────────────────────────────────────────
    st.divider()

    if errors:
        with st.expander(f"⚠️ {len(errors)} error(s) occurred", expanded=False):
            for e in errors:
                st.markdown(e)

    if not all_frames:
        st.warning("No data was collected from any company.")
        return

    master_df = pd.concat(all_frames, ignore_index=True)

    # Apply Consol Mapping Sheet lookup
    master_df = apply_mapping(master_df)

    # Apply Forex Rate based on Currency + Reporting Month + Item
    def _get_forex_rate(row):
        ccy = str(row.get("Currency", "")).strip().upper()
        item = str(row.get("Item", "")).strip()
        month = str(row.get("Reporting Month", "")).strip()

        # USD is always 1.0
        if ccy == "USD":
            return 1.0

        # Look up per-month rates for the currency
        month_rates = forex_rates.get(ccy, {})
        rates = month_rates.get(month, {"closing": 1.0, "average": 1.0})

        if item == "P&L":
            return rates["average"]
        else:
            return rates["closing"]

    master_df["Forex Rate"] = master_df.apply(_get_forex_rate, axis=1)
    master_df["Amount in USD (Reporting Currency)"] = (
        master_df["Transaction Value in Original Currency"] * master_df["Forex Rate"]
    )

    # Build Excel bytes and filename
    xlsx_bytes = to_excel_bytes(master_df)
    sd = datetime.strptime(start_date, "%Y-%m-%d")
    ed = datetime.strptime(end_date, "%Y-%m-%d")
    file_name = f"LXT_General_Ledger_{sd.strftime('%d%b%Y')}_to_{ed.strftime('%d%b%Y')}.xlsx"

    # Store in session state so the download button persists
    st.session_state["report_data"] = xlsx_bytes
    st.session_state["report_name"] = file_name
    st.session_state["report_rows"] = len(master_df)
    st.session_state["report_preview"] = master_df.head(100)

    # Store master_df for pivot report generation
    st.session_state["master_df"] = master_df


# ═══════════════════════════════════════════════════════════════
# Entry Point
# ═══════════════════════════════════════════════════════════════
if check_password():
    main_app()
