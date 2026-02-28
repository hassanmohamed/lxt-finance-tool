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
    # NOTE: "Foreign Debit"/"Foreign Credit" intentionally excluded —
    #   they contain foreign currency values (e.g. USD for a CAD company).
    #   We always want the native/home currency amounts.
    "Nat Debit": "Debit",
    "Nat Credit": "Credit",
}

QB_REPORT_COLUMNS = "account_name,tx_date,memo,name,txn_type,cust_name,vend_name,doc_num,subt_nat_amount,debt_amt,credit_amt,klass_name"


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

    # Drop foreign-currency columns so they never overwrite native amounts
    foreign_cols = [c for c in df.columns if c.startswith("Foreign")]
    if foreign_cols:
        df = df.drop(columns=foreign_cols)

    # Rename using map
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


def build_pivot_section(
    df: pd.DataFrame,
    month_keys: list[str],
    month_labels: list[str],
    section_label: str,
) -> list[dict]:
    """
    Build one pivot section (Consolidated / per-entity / per-CostCenter).

    Returns a list of row dicts with keys:
      'Description', month_labels[0], month_labels[1], month_labels[2], 'Variance', '_style'
    '_style' is 'section' | 'group' | 'detail' | 'total' for formatting.
    """
    import calendar

    rows: list[dict] = []

    # Section header row
    header = {"Description": section_label, "_style": "section"}
    for lbl in month_labels:
        header[lbl] = ""
    header["Variance"] = ""
    rows.append(header)

    # Get unique Statements preserving order
    statements = df["Statement"].dropna().unique()

    for stmt in statements:
        stmt_str = str(stmt).strip()
        if not stmt_str or stmt_str.lower() == "nan":
            continue

        stmt_df = df[df["Statement"] == stmt]

        # Group header
        group_row = {"Description": stmt_str, "_style": "group"}
        for lbl in month_labels:
            group_row[lbl] = ""
        group_row["Variance"] = ""
        rows.append(group_row)

        # Mapping detail lines
        mappings = stmt_df["Mapping"].dropna().unique()
        stmt_totals = {lbl: 0.0 for lbl in month_labels}

        for mapping in mappings:
            mapping_str = str(mapping).strip()
            if not mapping_str or mapping_str.lower() == "nan":
                continue

            detail = {"Description": f"  {mapping_str}", "_style": "detail"}

            for i, mk in enumerate(month_keys):
                lbl = month_labels[i]
                mask = (
                    (stmt_df["Mapping"] == mapping)
                    & (stmt_df["Reporting Month"] == mk)
                )
                val = stmt_df.loc[mask, "Amount in USD (Reporting Currency)"].sum()
                detail[lbl] = round(val, 2)
                stmt_totals[lbl] += val

            # Variance = latest month - previous month
            detail["Variance"] = round(
                detail[month_labels[0]] - detail[month_labels[1]], 2
            )
            rows.append(detail)

        # Statement total row
        total_row = {"Description": f"Total {stmt_str}", "_style": "total"}
        for lbl in month_labels:
            total_row[lbl] = round(stmt_totals[lbl], 2)
        total_row["Variance"] = round(
            stmt_totals[month_labels[0]] - stmt_totals[month_labels[1]], 2
        )
        rows.append(total_row)

    # Add an empty spacer row after each section
    spacer = {"Description": "", "_style": "detail"}
    for lbl in month_labels:
        spacer[lbl] = ""
    spacer["Variance"] = ""
    rows.append(spacer)

    return rows


def build_pivot_report(
    master_df: pd.DataFrame,
    selected_year: int,
    selected_month: int,
) -> tuple[pd.DataFrame, list[dict]]:
    """
    Build the full pivot P&L report from the master GL data.

    Returns (display_df, raw_rows) where raw_rows contains '_style' metadata.
    """
    import calendar

    # Determine 3 consecutive months (latest first)
    months = _prev_months(selected_year, selected_month, 3)
    month_keys = [_month_key(y, m) for y, m in months]
    month_labels = [calendar.month_abbr[m] + f" {y}" for y, m in months]

    # Filter to P&L items only (the pivot is an income statement)
    pl_df = master_df.copy()

    all_rows: list[dict] = []

    # ── Section 1: Consolidated (all countries) ──
    all_rows.extend(
        build_pivot_section(pl_df, month_keys, month_labels, "📊 CONSOLIDATED (All Countries)")
    )

    # ── Section 2: Per Legal Entity ──
    entities = sorted(pl_df["Company Country"].dropna().unique())
    for entity in entities:
        entity_df = pl_df[pl_df["Company Country"] == entity]
        all_rows.extend(
            build_pivot_section(entity_df, month_keys, month_labels, f"🏢 {entity}")
        )

    # ── Section 3: Per CostCenter ──
    cost_centers = pl_df["CostCenter"].dropna().astype(str).str.strip()
    cost_centers = sorted(cost_centers[cost_centers != ""].unique())
    for cc in cost_centers:
        cc_df = pl_df[pl_df["CostCenter"].astype(str).str.strip() == cc]
        all_rows.extend(
            build_pivot_section(cc_df, month_keys, month_labels, f"📁 CostCenter: {cc}")
        )

    # Build display DataFrame (without _style)
    columns = ["Description"] + month_labels + ["Variance"]
    display_df = pd.DataFrame(all_rows)[columns]

    return display_df, all_rows


def pivot_to_excel_bytes(rows: list[dict], columns: list[str]) -> bytes:
    """Write the pivot report to a styled Excel file."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, numbers

    wb = Workbook()
    ws = wb.active
    ws.title = "Pivot P&L Report"

    # Styles
    section_font = Font(bold=True, size=13, color="FFFFFF")
    section_fill = PatternFill(start_color="2D2D44", end_color="2D2D44", fill_type="solid")
    group_font = Font(bold=True, size=11, color="1B3A5C")
    group_fill = PatternFill(start_color="E8EEF4", end_color="E8EEF4", fill_type="solid")
    total_font = Font(bold=True, size=10)
    total_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    header_font = Font(bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="3A3A5C", end_color="3A3A5C", fill_type="solid")
    num_fmt = '#,##0.00'

    # Write header row
    for col_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    # Write data rows
    for row_idx, row_data in enumerate(rows, start=2):
        style = row_data.get("_style", "detail")
        for col_idx, col_name in enumerate(columns, start=1):
            val = row_data.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)

            if style == "section":
                cell.font = section_font
                cell.fill = section_fill
            elif style == "group":
                cell.font = group_font
                cell.fill = group_fill
            elif style == "total":
                cell.font = total_font
                cell.fill = total_fill

            # Number formatting for value columns
            if col_idx > 1 and isinstance(val, (int, float)):
                cell.number_format = num_fmt
                cell.alignment = Alignment(horizontal="right")

    # Set column widths
    ws.column_dimensions["A"].width = 40
    for col_idx in range(2, len(columns) + 1):
        from openpyxl.utils import get_column_letter
        ws.column_dimensions[get_column_letter(col_idx)].width = 18

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

    # ── Forex Rate Inputs ────────────────────────────────────
    with st.expander("💱 Forex Rates (per currency)", expanded=False):
        st.caption(
            "Enter the **Closing Rate** (for Balance Sheet items) and "
            "**Average Rate** (for P&L items) for each currency. "
            "USD is always 1.0."
        )
        forex_rates: dict[str, dict[str, float]] = {}
        # Default rates from Jan 2026 (ClosingRate2 / AverageRate2)
        default_rates = {
            "EGP": {"closing": 0.021308, "average": 0.021170},
            "CAD": {"closing": 0.734376, "average": 0.726164},
            "AUD": {"closing": 0.696136, "average": 0.679199},
            "RON": {"closing": 0.232580, "average": 0.230658},
            "INR": {"closing": 0.010906, "average": 0.011009},
            "EUR": {"closing": 1.184975, "average": 1.174584},
            "GBP": {"closing": 1.368925, "average": 1.353443},
        }
        for ccy in FOREX_CURRENCIES:
            defaults = default_rates.get(ccy, {"closing": 1.0, "average": 1.0})
            c1, c2, c3 = st.columns([1, 1, 1])
            with c1:
                st.markdown(f"**{ccy}**")
            with c2:
                closing = st.number_input(
                    f"{ccy} Closing Rate (B.S)",
                    min_value=0.0,
                    value=defaults["closing"],
                    step=0.0001,
                    format="%.6f",
                    key=f"fx_closing_{ccy}",
                    label_visibility="collapsed",
                )
            with c3:
                average = st.number_input(
                    f"{ccy} Average Rate (P&L)",
                    min_value=0.0,
                    value=defaults["average"],
                    step=0.0001,
                    format="%.6f",
                    key=f"fx_average_{ccy}",
                    label_visibility="collapsed",
                )
            forex_rates[ccy] = {"closing": closing, "average": average}
        # USD is always 1
        forex_rates["USD"] = {"closing": 1.0, "average": 1.0}

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
                display_df, raw_rows = build_pivot_report(
                    st.session_state["master_df"],
                    int(pivot_year),
                    int(pivot_month),
                )
                pivot_columns = list(display_df.columns)
                pivot_xlsx = pivot_to_excel_bytes(raw_rows, pivot_columns)

                m_labels = [
                    calendar.month_abbr[m] + f"{y}"
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
                    use_container_width=True,
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

                # Auth
                token_info = refresh_access_token(
                    client_id, client_secret, refresh_token
                )

                # Track the new refresh token
                new_refresh = token_info["refresh_token"]
                updated_tokens[key] = new_refresh

                # Also save locally (best-effort)
                if new_refresh != refresh_token:
                    _save_refresh_token(refresh_token, new_refresh)

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

    # Apply Forex Rate based on Item (P&L → average, B.S → closing)
    def _get_forex_rate(row):
        ccy = row.get("Currency", "")
        item = row.get("Item", "")
        rates = forex_rates.get(ccy, {"closing": 1.0, "average": 1.0})
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
