#!/usr/bin/env python3
"""
QuickBooks Multi-Company ETL Pipeline
======================================
Extracts General Ledger data from 9 QuickBooks companies,
transforms each with Pandas, and exports one consolidated Excel file.

Usage:
    1. Fill in .env with your QuickBooks credentials.
    2. pip install -r requirements.txt
    3. python quickbooks_etl.py [--start-date YYYY-MM-DD] [--end-date YYYY-MM-DD] [--output FILE.xlsx]
"""

import argparse
import base64
import json
import logging
import os
import sys
from datetime import datetime, timedelta

import pandas as pd
import requests
from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
QB_TOKEN_URL = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"
QB_BASE_URL = "https://quickbooks.api.intuit.com"
QB_SANDBOX_BASE_URL = "https://sandbox-quickbooks.api.intuit.com"

# The 9 companies: (display_label, realm_id_env_key, refresh_token_env_key)
COMPANIES = [
    ("LXT Egypt",           "LXT_EGYPT_REALM_ID",           "LXT_EGYPT_REFRESH_TOKEN"),
    ("LXT Canada",          "LXT_CANADA_REALM_ID",          "LXT_CANADA_REFRESH_TOKEN"),
    ("LXT Australia",       "LXT_AUSTRALIA_REALM_ID",       "LXT_AUSTRALIA_REFRESH_TOKEN"),
    ("LXT Romania",         "LXT_ROMANIA_REALM_ID",         "LXT_ROMANIA_REFRESH_TOKEN"),
    ("LXT India",           "LXT_INDIA_REALM_ID",           "LXT_INDIA_REFRESH_TOKEN"),
    ("LXT Germany",         "LXT_GERMANY_REALM_ID",         "LXT_GERMANY_REFRESH_TOKEN"),
    ("LXT UK",              "LXT_UK_REALM_ID",              "LXT_UK_REFRESH_TOKEN"),
    ("LXT USA",             "LXT_USA_REALM_ID",             "LXT_USA_REFRESH_TOKEN"),
    ("LXT Clickworker USA", "LXT_CLICKWORKER_USA_REALM_ID", "LXT_CLICKWORKER_USA_REFRESH_TOKEN"),
]

# Final output columns (in order).
OUTPUT_COLUMNS = [
    "Distribution account",
    "Date",
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
    "Company Country",
]

# Mapping: QuickBooks General Ledger column keys → output column names.
# The `columns` API parameter uses these internal keys, but the response
# ColTitle may return display names instead. We map BOTH so either works.
QB_COLUMN_MAP = {
    # Internal API keys (used in `columns` param)
    "account_name": "Distribution account",
    "tx_date": "Date",
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
    # Actual display-name ColTitles returned by the API
    "Account": "Distribution account",
    "Distribution Account": "Distribution account",
    "Transaction Type": "Transaction id",
    "Trans #": "Transaction id",
    "No.": "Number",
    "Num": "Number",
    "Customer": "Customer full name",
    "Vendor": "Supplier",
    "Memo/Description": "Memo/Description",
    "Class": "Class full name",
    "Amount": "Balance",
    "Debit": "Debit",
    "Credit": "Credit",
    # For multi-currency companies the API labels these as "Foreign"
    # but the values are in the company's home currency.
    "Foreign Debit": "Debit",
    "Foreign Credit": "Credit",
    "Nat Debit": "Debit",
    "Nat Credit": "Credit",
}

# Columns to request from the API (internal keys only).
QB_REPORT_COLUMNS = "account_name,tx_date,memo,name,txn_type,cust_name,vend_name,doc_num,subt_nat_amount,subt_nat_home_amount,debt_amt,credit_amt,klass_name"


# ═══════════════════════════════════════════════════════════════════════════
# Phase 1 — Authentication
# ═══════════════════════════════════════════════════════════════════════════
def refresh_access_token(
    client_id: str,
    client_secret: str,
    refresh_token: str,
) -> dict:
    """
    Exchange a refresh token for a new access token via
    QuickBooks OAuth 2.0 token endpoint.

    Returns dict with: access_token, refresh_token, expires_in.
    """
    credentials = f"{client_id}:{client_secret}"
    auth_header = base64.b64encode(credentials.encode()).decode()

    headers = {
        "Accept": "application/json",
        "Authorization": f"Basic {auth_header}",
        "Content-Type": "application/x-www-form-urlencoded",
    }
    payload = {
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
    }

    response = requests.post(QB_TOKEN_URL, headers=headers, data=payload, timeout=30)

    if response.status_code != 200:
        logger.error(
            "Token refresh failed (%s): %s", response.status_code, response.text
        )
        raise RuntimeError(
            f"Authentication failed (HTTP {response.status_code}). "
            "Check your credentials."
        )

    token_data = response.json()
    logger.info("Access token obtained (expires in %s s).", token_data.get("expires_in"))

    return {
        "access_token": token_data["access_token"],
        "refresh_token": token_data.get("refresh_token", refresh_token),
        "expires_in": token_data.get("expires_in"),
    }


# ═══════════════════════════════════════════════════════════════════════════
# Phase 2 — Extraction (General Ledger)
# ═══════════════════════════════════════════════════════════════════════════
def extract_general_ledger(
    access_token: str,
    realm_id: str,
    start_date: str,
    end_date: str,
    company_label: str,
    *,
    sandbox: bool = False,
    debug: bool = False,
) -> list[dict]:
    """
    Fetch the General Ledger report from QuickBooks and flatten
    the deeply nested JSON into a list of row-dicts.
    """
    base = QB_SANDBOX_BASE_URL if sandbox else QB_BASE_URL
    url = f"{base}/v3/company/{realm_id}/reports/GeneralLedger"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
    }
    params = {
        "start_date": start_date,
        "end_date": end_date,
        "columns": QB_REPORT_COLUMNS,
        "accounting_method": "Accrual",
    }

    logger.info("Fetching General Ledger (%s → %s) …", start_date, end_date)
    response = requests.get(url, headers=headers, params=params, timeout=120)

    if response.status_code != 200:
        logger.error(
            "API request failed (%s): %s", response.status_code, response.text
        )
        raise RuntimeError(
            f"Failed to fetch General Ledger (HTTP {response.status_code})."
        )

    report_json = response.json()

    # ── Debug: dump raw JSON to file ──────────────────────────────────────
    if debug:
        safe_name = company_label.replace(" ", "_").lower()
        debug_file = f"debug_raw_{safe_name}.json"
        with open(debug_file, "w") as f:
            json.dump(report_json, f, indent=2)
        logger.info("[DEBUG] Raw API response saved to %s", debug_file)

    return _parse_report_rows(report_json, company_label, debug=debug)


def _parse_report_rows(
    report_json: dict, company_label: str, *, debug: bool = False
) -> list[dict]:
    """
    Walk the nested QBO report JSON and return a flat list of dicts.

    Structure:
        Columns → Column[]  (header definitions)
        Rows    → Row[]
          Row can be:
            type "Data"    → ColData[] (actual line item)
            type "Section" → Header + nested Rows + Summary
    """
    columns_meta = report_json.get("Columns", {}).get("Column", [])
    col_keys = [col.get("ColTitle", "").strip() for col in columns_meta]

    if debug:
        logger.info("[DEBUG] [%s] Column titles from API: %s", company_label, col_keys)
        # Also log the raw Column metadata for full visibility
        logger.info(
            "[DEBUG] [%s] Column metadata:\n%s",
            company_label,
            json.dumps(columns_meta, indent=2),
        )

    rows_section = report_json.get("Rows", {}).get("Row", [])
    flat_rows: list[dict] = []
    _walk_rows(rows_section, col_keys, flat_rows)

    if debug and flat_rows:
        logger.info(
            "[DEBUG] [%s] First 3 raw parsed rows:\n%s",
            company_label,
            json.dumps(flat_rows[:3], indent=2),
        )

    logger.info("[%s] Parsed %d data rows from report.", company_label, len(flat_rows))
    return flat_rows


def _walk_rows(rows: list, col_keys: list[str], accumulator: list[dict]) -> None:
    """Recursively walk Row / Section nodes and collect only Data rows."""
    for row in rows:
        row_type = row.get("type", "Data")

        if row_type == "Data":
            col_data = row.get("ColData", [])
            record = {}
            for idx, cell in enumerate(col_data):
                key = col_keys[idx] if idx < len(col_keys) else f"col_{idx}"
                record[key] = cell.get("value", "")
            accumulator.append(record)

        elif row_type == "Section":
            # Skip section headers — they contain account names only,
            # not transaction-level data. We only want nested Data rows.
            nested = row.get("Rows", {}).get("Row", [])
            if nested:
                _walk_rows(nested, col_keys, accumulator)

            # Skip Summary rows (totals) — not transaction-level data


# ═══════════════════════════════════════════════════════════════════════════
# Phase 3 — Transformation
# ═══════════════════════════════════════════════════════════════════════════
def transform(
    raw_rows: list[dict], company_label: str, *, debug: bool = False
) -> pd.DataFrame:
    """
    Convert raw row-dicts into a cleaned DataFrame with the exact
    required columns, filtering, and a Company Country label.
    """
    if not raw_rows:
        logger.warning("No data rows to transform for %s.", company_label)
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    df = pd.DataFrame(raw_rows)

    if debug:
        logger.info(
            "[DEBUG] [%s] DataFrame columns BEFORE rename: %s",
            company_label,
            list(df.columns),
        )
        logger.info(
            "[DEBUG] [%s] DataFrame shape: %d rows × %d cols",
            company_label,
            len(df),
            len(df.columns),
        )
        logger.info(
            "[DEBUG] [%s] First 5 rows BEFORE rename:\n%s",
            company_label,
            df.head().to_string(),
        )

    # ── Rename columns ────────────────────────────────────────────────────
    # Build rename map from whichever keys are actually present in the data
    rename_map = {}
    for raw_key, output_name in QB_COLUMN_MAP.items():
        if raw_key in df.columns and raw_key != output_name:
            rename_map[raw_key] = output_name
    df = df.rename(columns=rename_map)

    if debug:
        logger.info(
            "[DEBUG] [%s] Rename map applied: %s", company_label, rename_map
        )
        logger.info(
            "[DEBUG] [%s] DataFrame columns AFTER rename: %s",
            company_label,
            list(df.columns),
        )

    # ── Enrich Transaction id (type + doc number) ─────────────────────────
    if "Transaction id" in df.columns and "Number" in df.columns:
        df["Transaction id"] = (
            df["Transaction id"].astype(str).str.strip()
            + " #"
            + df["Number"].astype(str).str.strip()
        )

    # ── Ensure all required columns exist ─────────────────────────────────
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    # ── Add Company Country ───────────────────────────────────────────────
    df["Company Country"] = company_label

    # Keep only required columns, in order
    df = df[OUTPUT_COLUMNS]

    # ── Convert numeric columns ───────────────────────────────────────────
    for num_col in ("Balance", "Debit", "Credit"):
        df[num_col] = pd.to_numeric(df[num_col], errors="coerce")

    # ── Filter: drop rows where Distribution account is empty / null ──────
    df["Distribution account"] = df["Distribution account"].astype(str).str.strip()
    before_filter = len(df)
    df = df[
        (df["Distribution account"] != "")
        & (df["Distribution account"].str.lower() != "none")
        & (df["Distribution account"].str.lower() != "nan")
    ]
    df = df.dropna(subset=["Distribution account"])
    df = df.reset_index(drop=True)

    if debug:
        logger.info(
            "[DEBUG] [%s] Filter removed %d rows (before: %d, after: %d)",
            company_label,
            before_filter - len(df),
            before_filter,
            len(df),
        )

    logger.info(
        "[%s] %d rows after transformation & filtering.", company_label, len(df)
    )
    return df


# ═══════════════════════════════════════════════════════════════════════════
# Phase 4 — Export
# ═══════════════════════════════════════════════════════════════════════════
def export_to_excel(df: pd.DataFrame, output_path: str) -> None:
    """Write the consolidated DataFrame to a clean .xlsx file."""
    df.to_excel(output_path, index=False)
    logger.info("Exported %d rows to %s", len(df), output_path)


# ═══════════════════════════════════════════════════════════════════════════
# CLI / Main
# ═══════════════════════════════════════════════════════════════════════════
def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="QuickBooks Multi-Company General Ledger → Excel ETL pipeline.",
    )
    parser.add_argument(
        "--start-date",
        default=(datetime.now() - timedelta(days=365)).strftime("%Y-%m-%d"),
        help="Start date for the report (YYYY-MM-DD). Default: 1 year ago.",
    )
    parser.add_argument(
        "--end-date",
        default=datetime.now().strftime("%Y-%m-%d"),
        help="End date for the report (YYYY-MM-DD). Default: today.",
    )
    parser.add_argument(
        "--output",
        default="quickbooks_consolidated_report.xlsx",
        help="Output Excel file path. Default: quickbooks_consolidated_report.xlsx",
    )
    parser.add_argument(
        "--sandbox",
        action="store_true",
        help="Use the QuickBooks sandbox environment instead of production.",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable debug mode: dumps raw API JSON to files and logs column info.",
    )
    parser.add_argument(
        "--company",
        type=str,
        default=None,
        help="Run for a single company only (e.g. 'LXT Romania'). Default: all 9.",
    )
    return parser.parse_args()


def main() -> None:
    """Run the full multi-company ETL pipeline."""
    load_dotenv()
    args = _parse_args()

    # ── Load shared credentials ───────────────────────────────────────────
    client_id = os.getenv("QB_CLIENT_ID_DEV") or os.getenv("QB_CLIENT_ID")
    client_secret = os.getenv("QB_CLIENT_SECRET_DEV") or os.getenv("QB_CLIENT_SECRET")

    if not client_id or not client_secret:
        logger.error(
            "Missing QB_CLIENT_ID / QB_CLIENT_SECRET in .env. "
            "Set QB_CLIENT_ID_DEV + QB_CLIENT_SECRET_DEV or "
            "QB_CLIENT_ID + QB_CLIENT_SECRET."
        )
        sys.exit(1)

    # ── Determine which companies to process ──────────────────────────────
    companies_to_run = COMPANIES
    if args.company:
        companies_to_run = [
            c for c in COMPANIES if c[0].lower() == args.company.lower()
        ]
        if not companies_to_run:
            logger.error(
                "Company '%s' not found. Available: %s",
                args.company,
                ", ".join(c[0] for c in COMPANIES),
            )
            sys.exit(1)

    # ── Loop over companies ───────────────────────────────────────────────
    all_frames: list[pd.DataFrame] = []
    success_count = 0
    fail_count = 0

    for company_label, realm_key, token_key in companies_to_run:
        realm_id = os.getenv(realm_key)
        refresh_token = os.getenv(token_key)

        if not realm_id or not refresh_token:
            logger.warning(
                "[%s] Skipping — missing env vars: %s / %s",
                company_label,
                realm_key,
                token_key,
            )
            fail_count += 1
            continue

        try:
            # Authenticate
            logger.info("═" * 60)
            logger.info("Processing: %s", company_label)
            logger.info("═" * 60)
            token_info = refresh_access_token(client_id, client_secret, refresh_token)
            access_token = token_info["access_token"]

            # Log new refresh token if it changed
            new_refresh = token_info["refresh_token"]
            if new_refresh != refresh_token:
                logger.info(
                    "[%s] New refresh token issued. Update %s in .env:\n  %s",
                    company_label,
                    token_key,
                    new_refresh,
                )

            # Extract
            raw_rows = extract_general_ledger(
                access_token,
                realm_id,
                start_date=args.start_date,
                end_date=args.end_date,
                company_label=company_label,
                sandbox=args.sandbox,
                debug=args.debug,
            )

            # Transform
            df = transform(raw_rows, company_label, debug=args.debug)
            if not df.empty:
                all_frames.append(df)

            success_count += 1

        except Exception as exc:
            logger.error("[%s] Failed: %s", company_label, exc)
            fail_count += 1
            continue

    # ── Consolidate & Export ───────────────────────────────────────────────
    logger.info("═" * 60)
    logger.info(
        "Completed: %d succeeded, %d failed out of %d companies.",
        success_count,
        fail_count,
        len(companies_to_run),
    )

    if not all_frames:
        logger.warning("No data collected from any company. No file created.")
        sys.exit(0)

    master_df = pd.concat(all_frames, ignore_index=True)
    export_to_excel(master_df, args.output)

    logger.info("✅  ETL pipeline completed successfully.")


if __name__ == "__main__":
    main()
