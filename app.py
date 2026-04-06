"""
Month-End Close Agent — Streamlit Application
Finance Department
Diagnostic reporting for QuickBooks month-end close.

Reports:
  1. Flux (Variance) Narrative
  2. Recurring Vendor Gap Analysis ("Missing Bill")
  3. Suspense & Misc Resolution Worksheet
  4. Materiality & Risk Threshold
  5. IIF Import Pre-Flight Validation
  6. Multi-Source Reconciliation Summary
"""

import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import os
import subprocess
import tempfile
import json
import datetime
import warnings
from collections import defaultdict

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────
# 0 · APPLE-BRANDED THEME
# ─────────────────────────────────────────────────────────────
APPLE = {
    "blue":        "#0071E3",
    "deep_blue":   "#0055CC",
    "bright_blue": "#1A8CFF",
    "near_black":  "#1D1D1F",
    "gray_sec":    "#6E6E73",
    "sys_gray":    "#8E8E93",
    "mid_gray":    "#D2D2D7",
    "light_gray":  "#F5F5F7",
    "white":       "#FFFFFF",
    "red":         "#FF3B30",
    "green":       "#34C759",
    "orange":      "#FF9500",
    "yellow":      "#FFCC00",
}

st.set_page_config(
    page_title="Month-End Close Agent · Month-End Close",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(f"""
<style>
    /* ── Global ── */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] {{
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Helvetica Neue', sans-serif;
        color: {APPLE["near_black"]};
    }}
    /* ── Header banner ── */
    .apple-header {{
        background: linear-gradient(135deg, {APPLE["bright_blue"]}, {APPLE["deep_blue"]});
        border-radius: 16px;
        padding: 2rem 2.5rem;
        margin-bottom: 1.5rem;
        color: {APPLE["white"]};
    }}
    .apple-header h1 {{
        font-weight: 700; font-size: 1.75rem; margin: 0; letter-spacing: -0.02em;
    }}
    .apple-header p {{
        font-weight: 400; font-size: 0.95rem; opacity: 0.85; margin: 0.35rem 0 0 0;
    }}
    /* ── Cards ── */
    .metric-card {{
        background: {APPLE["light_gray"]};
        border: 1px solid {APPLE["mid_gray"]};
        border-radius: 12px;
        padding: 1.15rem 1.35rem;
        margin-bottom: 0.75rem;
    }}
    .metric-card .label {{
        font-size: 0.75rem; font-weight: 500; color: {APPLE["gray_sec"]};
        text-transform: uppercase; letter-spacing: 0.04em;
    }}
    .metric-card .value {{
        font-size: 1.5rem; font-weight: 700; color: {APPLE["near_black"]};
    }}
    /* ── Status pills ── */
    .pill-pass {{
        background: {APPLE["green"]}22; color: {APPLE["green"]};
        padding: 2px 10px; border-radius: 20px; font-weight: 600; font-size: 0.8rem;
    }}
    .pill-fail {{
        background: {APPLE["red"]}22; color: {APPLE["red"]};
        padding: 2px 10px; border-radius: 20px; font-weight: 600; font-size: 0.8rem;
    }}
    .pill-warn {{
        background: {APPLE["orange"]}22; color: {APPLE["orange"]};
        padding: 2px 10px; border-radius: 20px; font-weight: 600; font-size: 0.8rem;
    }}
    /* ── Section headers ── */
    .section-head {{
        font-weight: 700; font-size: 1.05rem; color: {APPLE["near_black"]};
        border-bottom: 2px solid {APPLE["blue"]};
        padding-bottom: 0.4rem; margin: 1.5rem 0 0.75rem 0;
        text-transform: uppercase; letter-spacing: 0.03em;
    }}
    /* ── Table tweaks ── */
    .stDataFrame thead th {{
        background: {APPLE["light_gray"]} !important;
        font-weight: 600 !important;
        color: {APPLE["near_black"]} !important;
    }}
    /* ── Sidebar ── */
    section[data-testid="stSidebar"] {{
        background: {APPLE["light_gray"]};
    }}
    section[data-testid="stSidebar"] .stMarkdown p {{
        color: {APPLE["gray_sec"]};
        font-size: 0.85rem;
    }}
    /* ── Narrative blocks ── */
    .narrative-block {{
        background: {APPLE["white"]};
        border-left: 4px solid {APPLE["blue"]};
        padding: 1rem 1.25rem;
        border-radius: 0 8px 8px 0;
        margin-bottom: 0.75rem;
        font-size: 0.92rem;
        line-height: 1.6;
    }}
    /* ── Footer ── */
    .apple-footer {{
        text-align: center;
        color: {APPLE["sys_gray"]};
        font-size: 0.75rem;
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid {APPLE["mid_gray"]};
    }}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# 1 · HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────

def metric_card(label: str, value: str) -> str:
    return f"""<div class="metric-card">
        <div class="label">{label}</div>
        <div class="value">{value}</div>
    </div>"""


def pill(status: str, text: str) -> str:
    cls = {"pass": "pill-pass", "fail": "pill-fail", "warn": "pill-warn"}.get(status, "pill-warn")
    return f'<span class="{cls}">{text}</span>'


def section(title: str):
    st.markdown(f'<div class="section-head">{title}</div>', unsafe_allow_html=True)


def narrative(text: str):
    st.markdown(f'<div class="narrative-block">{text}</div>', unsafe_allow_html=True)


def fmt_currency(v):
    try:
        return f"${abs(float(v)):,.2f}"
    except (ValueError, TypeError):
        return "$0.00"


# ─────────────────────────────────────────────────────────────
# 2 · DATA PARSERS
# ─────────────────────────────────────────────────────────────

@st.cache_data
def parse_gl(file) -> pd.DataFrame:
    """Parse a General Ledger from CSV or QuickBooks Excel export.

    Sniffs the actual file content to detect Excel binaries even when
    the extension is .csv (common with QuickBooks exports).
    """
    # Read first 4 bytes to detect file type
    header = file.read(4)
    file.seek(0)
    is_excel = header[:2] == b"PK" or header[:8] == b"\xd0\xcf\x11\xe0"

    name = file.name.lower()
    if is_excel or name.endswith((".xlsx", ".xls")):
        return _parse_gl_excel(file)
    else:
        return _parse_gl_csv(file)


def _parse_gl_csv(file) -> pd.DataFrame:
    """Handles standard CSV GL exports with named columns."""
    # Try multiple encodings for robustness
    for encoding in ["utf-8", "latin-1", "cp1252"]:
        try:
            file.seek(0)
            df = pd.read_csv(file, encoding=encoding)
            break
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue
    else:
        file.seek(0)
        df = pd.read_csv(file, encoding="latin-1", on_bad_lines="skip")
    df.columns = [c.strip() for c in df.columns]
    col_map = {}
    for c in df.columns:
        cl = c.lower()
        if "date" in cl:
            col_map["Date"] = c
        elif "account" in cl and "code" in cl:
            col_map["Account"] = c
        elif "account" in cl and "Account" not in col_map:
            col_map["Account"] = c
        elif "name" in cl or "vendor" in cl:
            col_map["Name"] = c
        elif "memo" in cl or "desc" in cl:
            col_map["Memo"] = c
        elif "debit" in cl:
            col_map["Debit"] = c
        elif "credit" in cl:
            col_map["Credit"] = c
        elif "type" in cl:
            col_map["Type"] = c
        elif "split" in cl:
            col_map["Split"] = c
        elif "num" in cl:
            col_map["Num"] = c
    rename = {v: k for k, v in col_map.items()}
    df = df.rename(columns=rename)
    return _clean_gl(df)


def _parse_gl_excel(file) -> pd.DataFrame:
    """Parse QuickBooks Desktop hierarchical Excel GL export."""
    xls = pd.ExcelFile(file, engine="openpyxl")
    sheets = xls.sheet_names
    target = "Sheet1" if "Sheet1" in sheets else sheets[-1]
    raw = pd.read_excel(file, sheet_name=target, engine="openpyxl", header=None)

    # Detect header row — look for a row containing "Type" and "Date"
    header_idx = None
    for i in range(min(10, len(raw))):
        row_vals = raw.iloc[i].astype(str).str.strip().str.lower().tolist()
        if "type" in row_vals and "date" in row_vals:
            header_idx = i
            break
    if header_idx is None:
        header_idx = 0

    col_indices = {}
    known_header_names = {"type", "date", "num", "name", "memo", "split", "debit", "credit", "balance"}
    for idx in range(raw.shape[1]):
        val = raw.iloc[header_idx, idx]
        if pd.isna(val):
            continue
        vl = str(val).lower().strip()
        if vl in known_header_names:
            col_indices[vl.capitalize()] = idx

    # Determine account columns: columns BEFORE the first known field column.
    # In QuickBooks exports, account hierarchy is in the leftmost columns,
    # and data fields (Type, Date, etc.) start after that.
    first_field_col = min(col_indices.values()) if col_indices else 0
    acct_cols = [c for c in range(first_field_col)]

    # Build a hierarchy tracker — deepest non-null column wins
    records = []
    # Track the most recent account value at each column level
    acct_stack = {}
    current_account = ""

    for i in range(header_idx + 1, len(raw)):
        row = raw.iloc[i]
        # Transaction rows have a date
        date_val = row.get(col_indices.get("Date"), None)

        # First, check if any account column has a value to update the hierarchy
        for ac in acct_cols:
            v = row.get(ac)
            if pd.notna(v):
                val_str = str(v).strip()
                if val_str and not val_str.lower().startswith("total"):
                    acct_stack[ac] = val_str
                    # Clear deeper levels when a higher level changes
                    for deeper in acct_cols:
                        if deeper > ac:
                            acct_stack.pop(deeper, None)

        if pd.notna(date_val) and str(date_val).strip() != "":
            # Use the deepest level in the hierarchy stack
            if acct_stack:
                deepest_col = max(acct_stack.keys())
                current_account = acct_stack[deepest_col]
            rec = {"Account": current_account}
            for field, cidx in col_indices.items():
                rec[field] = row.get(cidx)
            records.append(rec)

    df = pd.DataFrame(records)
    return _clean_gl(df)


def _clean_gl(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize GL DataFrame."""
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    for col in ["Debit", "Credit"]:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(r"[,$]", "", regex=True),
                errors="coerce",
            ).fillna(0.0)
    for col in ["Name", "Memo", "Account", "Type", "Split"]:
        if col in df.columns:
            df[col] = df[col].astype(str).fillna("")
    # Add helper columns
    df["Amount"] = df.get("Debit", 0) - df.get("Credit", 0)
    if "Date" in df.columns:
        df["YearMonth"] = df["Date"].dt.to_period("M")
    # Strip account number/name
    if "Account" in df.columns:
        df["AcctClean"] = df["Account"].str.replace(r"^\d+\s*·\s*", "", regex=True).str.strip()
    return df


@st.cache_data
def parse_iif(file) -> dict:
    """Parse a QuickBooks IIF file into accounts and classes."""
    content = file.read().decode("utf-8", errors="replace")
    lines = content.split("\n")
    accounts, classes = [], []
    section = None
    headers = []
    for line in lines:
        parts = line.rstrip("\r").split("\t")
        if not parts:
            continue
        tag = parts[0].strip()
        if tag == "!ACCNT":
            section = "ACCNT"
            headers = [p.strip() for p in parts]
        elif tag == "!CLASS":
            section = "CLASS"
            headers = [p.strip() for p in parts]
        elif tag == "ACCNT" and section == "ACCNT":
            row = dict(zip(headers, [p.strip() for p in parts]))
            accounts.append(row)
        elif tag == "CLASS" and section == "CLASS":
            row = dict(zip(headers, [p.strip() for p in parts]))
            classes.append(row)
        elif tag.startswith("!") and tag not in ("!HDR",):
            section = None

    acct_df = pd.DataFrame(accounts) if accounts else pd.DataFrame(columns=["NAME", "ACCNTTYPE", "ACCNUM", "HIDDEN"])
    class_df = pd.DataFrame(classes) if classes else pd.DataFrame(columns=["NAME", "HIDDEN"])
    return {
        "accounts": acct_df,
        "classes": class_df,
        "account_names": set(acct_df["NAME"].tolist()) if "NAME" in acct_df.columns else set(),
        "active_accounts": set(
            acct_df.loc[acct_df.get("HIDDEN", pd.Series(["N"] * len(acct_df))) == "N", "NAME"].tolist()
        ) if "NAME" in acct_df.columns else set(),
        "active_classes": set(
            class_df.loc[class_df.get("HIDDEN", pd.Series(["N"] * len(class_df))) == "N", "NAME"].tolist()
        ) if "NAME" in class_df.columns else set(),
    }


@st.cache_data
def parse_coa_csv(file) -> dict:
    """Parse a CSV-based Chart of Accounts."""
    df = pd.read_csv(file)
    df.columns = [c.strip() for c in df.columns]
    names = set()
    for c in df.columns:
        if "name" in c.lower() or "account" in c.lower():
            names.update(df[c].dropna().astype(str).tolist())
    return {
        "accounts": df,
        "classes": pd.DataFrame(),
        "account_names": names,
        "active_accounts": names,
        "active_classes": set(),
    }


def extract_pdf_text(file) -> str:
    """Best-effort text extraction from an uploaded PDF."""
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(file.read())) as pdf:
            return "\n".join(page.extract_text() or "" for page in pdf.pages)
    except ImportError:
        try:
            from PyPDF2 import PdfReader
            reader = PdfReader(io.BytesIO(file.read()))
            return "\n".join(page.extract_text() or "" for page in reader.pages)
        except Exception:
            return ""
    except Exception:
        return ""


# ─────────────────────────────────────────────────────────────
# 3 · REPORT ENGINES
# ─────────────────────────────────────────────────────────────

def report_flux(gl: pd.DataFrame):
    """Report 1 — Flux (Variance) Narrative."""
    section("Flux (Variance) Narrative Report")

    if "YearMonth" not in gl.columns or gl["YearMonth"].isna().all():
        st.warning("Date column could not be parsed — unable to compute month-over-month flux.")
        return

    periods = sorted(gl["YearMonth"].dropna().unique())
    if len(periods) < 2:
        st.info("Only one period detected in the GL. At least two months are required for variance analysis.")
        return

    curr = periods[-1]
    prev = periods[-2]
    st.markdown(f"**Comparing:** `{prev}` → `{curr}`")

    curr_df = gl[gl["YearMonth"] == curr]
    prev_df = gl[gl["YearMonth"] == prev]

    curr_totals = curr_df.groupby("Account")["Amount"].sum()
    prev_totals = prev_df.groupby("Account")["Amount"].sum()

    flux = pd.DataFrame({"Current": curr_totals, "Prior": prev_totals}).fillna(0)
    flux["Variance"] = flux["Current"] - flux["Prior"]
    flux["Var_%"] = np.where(flux["Prior"] != 0, (flux["Variance"] / flux["Prior"].abs()) * 100, np.nan)
    flux = flux.sort_values("Variance", key=abs, ascending=False)

    # Summary metrics
    c1, c2, c3 = st.columns(3)
    total_var = flux["Variance"].sum()
    top_driver = flux.index[0] if len(flux) > 0 else "N/A"
    large_moves = (flux["Variance"].abs() > flux["Variance"].abs().quantile(0.9)).sum()
    with c1:
        st.markdown(metric_card("Net Variance", f"${total_var:,.2f}"), unsafe_allow_html=True)
    with c2:
        st.markdown(metric_card("Top Driver", str(top_driver)[:40]), unsafe_allow_html=True)
    with c3:
        st.markdown(metric_card("Large Moves (>P90)", str(large_moves)), unsafe_allow_html=True)

    # Narrative generation — scan memos and vendors for top movers
    section("AI-Generated Narrative")
    top_n = flux.head(10)
    for acct, row in top_n.iterrows():
        direction = "increased" if row["Variance"] > 0 else "decreased"
        pct = f"{abs(row['Var_%']):.1f}%" if pd.notna(row["Var_%"]) else "N/A"
        # Scan memos for this account in current period
        acct_txns = curr_df[curr_df["Account"] == acct]
        memo_sample = acct_txns["Memo"].dropna().unique()[:3]
        vendor_sample = acct_txns["Name"].dropna().unique()[:3]
        memo_str = "; ".join(str(m) for m in memo_sample if str(m).strip() and str(m) != "nan") or "no memo detail"
        vendor_str = ", ".join(str(v) for v in vendor_sample if str(v).strip() and str(v) != "nan") or "various"

        # Detect patterns
        notes = []
        if acct_txns.shape[0] > 0:
            avg_txn = acct_txns["Amount"].abs().mean()
            max_txn = acct_txns["Amount"].abs().max()
            if max_txn > avg_txn * 3 and acct_txns.shape[0] > 1:
                notes.append("an unusually large single transaction was detected")
            dup_memos = acct_txns.groupby("Memo").size()
            dup_memos = dup_memos[dup_memos > 1]
            if len(dup_memos) > 0:
                notes.append(f"possible duplicate entries for: {', '.join(dup_memos.index[:2])}")

        note_str = ". Additionally, " + "; ".join(notes) + "." if notes else "."
        txt = (
            f"<strong>{acct}</strong> {direction} by <strong>${abs(row['Variance']):,.2f}</strong> "
            f"({pct}), moving from ${row['Prior']:,.2f} to ${row['Current']:,.2f}. "
            f"Key vendors: {vendor_str}. Recent memos reference: <em>{memo_str}</em>{note_str}"
        )
        narrative(txt)

    # Detail table
    section("Variance Detail")
    display = flux.copy()
    display.index.name = "Account"
    display = display.reset_index()
    for c in ["Current", "Prior", "Variance"]:
        display[c] = display[c].map(lambda v: f"${v:,.2f}")
    display["Var_%"] = display["Var_%"].map(lambda v: f"{v:.1f}%" if pd.notna(v) else "—")
    st.dataframe(display, use_container_width=True, hide_index=True)


def report_vendor_gap(gl: pd.DataFrame):
    """Report 2 — Recurring Vendor Gap Analysis ('Missing Bill')."""
    section("Recurring Vendor Gap Analysis")

    if "Name" not in gl.columns or "YearMonth" not in gl.columns:
        st.warning("GL must contain Name and Date columns for vendor gap analysis.")
        return

    vendors = gl[gl["Name"].str.strip() != "nan"].copy()
    if vendors.empty:
        st.info("No vendor names found in the GL.")
        return

    periods = sorted(vendors["YearMonth"].dropna().unique())
    if len(periods) < 3:
        st.info("At least 3 months of history are needed to detect recurring vendors.")
        return

    curr = periods[-1]
    history = vendors[vendors["YearMonth"] != curr]
    curr_vendors = set(vendors[vendors["YearMonth"] == curr]["Name"].unique())

    # Calculate vendor frequency
    vendor_periods = history.groupby("Name")["YearMonth"].nunique()
    total_hist_months = len(periods) - 1
    recurring = vendor_periods[vendor_periods >= max(2, total_hist_months * 0.5)]

    missing = [v for v in recurring.index if v not in curr_vendors and str(v).strip() and str(v) != "nan"]

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(metric_card("Recurring Vendors", str(len(recurring))), unsafe_allow_html=True)
    with c2:
        st.markdown(metric_card("Missing This Month", str(len(missing))), unsafe_allow_html=True)
    with c3:
        est_accrual = 0
        for v in missing:
            avg = history[history["Name"] == v]["Credit"].mean()
            est_accrual += avg if pd.notna(avg) else 0
        st.markdown(metric_card("Est. Accrual Total", f"${est_accrual:,.2f}"), unsafe_allow_html=True)

    if missing:
        section("Suggested Accruals")
        rows = []
        for v in missing:
            v_hist = history[history["Name"] == v]
            avg_amt = v_hist["Credit"].mean()
            last_date = v_hist["Date"].max()
            freq = v_hist["YearMonth"].nunique()
            typical_acct = v_hist["Account"].mode().iloc[0] if not v_hist["Account"].mode().empty else "Unknown"
            rows.append({
                "Vendor": v,
                "Historical Frequency": f"{freq} of {total_hist_months} months",
                "Avg Monthly Amount": f"${avg_amt:,.2f}" if pd.notna(avg_amt) else "—",
                "Last Seen": str(last_date.date()) if pd.notna(last_date) else "—",
                "Typical Account": typical_acct,
                "Suggested Accrual": f"${avg_amt:,.2f}" if pd.notna(avg_amt) else "—",
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    else:
        st.success("All recurring vendors have activity in the current period.")


def report_suspense(gl: pd.DataFrame, coa: dict | None):
    """Report 3 — Suspense & Misc Resolution Worksheet."""
    section("Suspense & Misc Resolution Worksheet")

    patterns = r"(?i)(suspense|misc|other|unclass|uncategoriz|clearing|unknown|unallocat)"
    mask = gl["Account"].str.contains(patterns, na=False) | gl["Memo"].str.contains(patterns, na=False)
    if "Split" in gl.columns:
        mask = mask | gl["Split"].str.contains(patterns, na=False)

    flagged = gl[mask].copy()

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(metric_card("Flagged Transactions", str(len(flagged))), unsafe_allow_html=True)
    with c2:
        total_amt = flagged["Amount"].abs().sum()
        st.markdown(metric_card("Total at Risk", f"${total_amt:,.2f}"), unsafe_allow_html=True)

    if flagged.empty:
        st.success("No transactions found in suspense, misc, or clearing accounts.")
        return

    # Suggest reclassifications using COA keyword matching
    suggestions = []
    active_accts = coa.get("active_accounts", set()) if coa else set()
    acct_list = sorted(active_accts)

    for _, txn in flagged.iterrows():
        memo = str(txn.get("Memo", "")).lower()
        name = str(txn.get("Name", "")).lower()
        search_text = memo + " " + name

        best_match = ""
        best_score = 0
        for acct in acct_list:
            acct_lower = acct.lower()
            # Split account name into keywords
            keywords = re.split(r"[:\s\-&/]+", acct_lower)
            keywords = [k for k in keywords if len(k) > 2]
            score = sum(1 for k in keywords if k in search_text)
            if score > best_score:
                best_score = score
                best_match = acct

        suggestions.append({
            "Date": str(txn.get("Date", ""))[:10],
            "Current Account": txn.get("Account", ""),
            "Name / Vendor": txn.get("Name", ""),
            "Memo": str(txn.get("Memo", ""))[:80],
            "Amount": f"${txn['Amount']:,.2f}",
            "Suggested Reclass": best_match if best_score >= 1 else "— manual review —",
            "Confidence": f"{'High' if best_score >= 3 else 'Medium' if best_score >= 1 else 'Low'}",
        })

    st.dataframe(pd.DataFrame(suggestions), use_container_width=True, hide_index=True)


def report_materiality(gl: pd.DataFrame, threshold: float):
    """Report 4 — Materiality & Risk Threshold."""
    section("Materiality & Risk Threshold Report")

    st.markdown(f"**Active Threshold:** ${threshold:,.0f}")

    # Flag categories: large uncategorized, suspense, large single txns, imbalances
    flags = []

    # Large transactions above threshold
    large = gl[gl["Amount"].abs() >= threshold].copy()
    for _, txn in large.iterrows():
        risk = "Low"
        reasons = []

        acct = str(txn.get("Account", "")).lower()
        memo = str(txn.get("Memo", "")).lower()

        if re.search(r"(suspense|misc|other|clearing|unknown)", acct):
            risk = "High"
            reasons.append("sitting in suspense/misc account")
        if memo in ("", "nan", "none"):
            risk = "High" if risk == "High" else "Medium"
            reasons.append("no memo/description")
        if abs(txn["Amount"]) >= threshold * 5:
            risk = "High"
            reasons.append("exceeds 5x materiality threshold")

        flags.append({
            "Date": str(txn.get("Date", ""))[:10],
            "Account": txn.get("Account", ""),
            "Name": txn.get("Name", ""),
            "Memo": str(txn.get("Memo", ""))[:60],
            "Amount": f"${txn['Amount']:,.2f}",
            "Risk": risk,
            "Reason": "; ".join(reasons) if reasons else "Material amount",
        })

    df_flags = pd.DataFrame(flags)
    if df_flags.empty:
        st.success(f"No transactions exceed the ${threshold:,.0f} materiality threshold.")
        return

    # Summary
    c1, c2, c3 = st.columns(3)
    high_ct = (df_flags["Risk"] == "High").sum()
    med_ct = (df_flags["Risk"] == "Medium").sum()
    with c1:
        st.markdown(metric_card("Flagged Items", str(len(df_flags))), unsafe_allow_html=True)
    with c2:
        st.markdown(metric_card("High Risk", f"{pill('fail', str(high_ct))}"), unsafe_allow_html=True)
    with c3:
        st.markdown(metric_card("Medium Risk", f"{pill('warn', str(med_ct))}"), unsafe_allow_html=True)

    st.dataframe(df_flags.sort_values("Risk", ascending=True), use_container_width=True, hide_index=True)


def report_iif_preflight(gl: pd.DataFrame, coa: dict | None):
    """Report 5 — IIF Import Pre-Flight Validation."""
    section("IIF Import Pre-Flight Validation")

    checks = []

    # Check 1: Debits = Credits per period
    if "YearMonth" in gl.columns:
        for period in sorted(gl["YearMonth"].dropna().unique()):
            p_df = gl[gl["YearMonth"] == period]
            total_dr = p_df["Debit"].sum()
            total_cr = p_df["Credit"].sum()
            diff = abs(total_dr - total_cr)
            balanced = diff < 0.01
            checks.append({
                "Check": f"Debits = Credits ({period})",
                "Status": "PASS" if balanced else "FAIL",
                "Detail": f"DR: ${total_dr:,.2f}  |  CR: ${total_cr:,.2f}  |  Diff: ${diff:,.2f}",
            })
    else:
        total_dr = gl["Debit"].sum()
        total_cr = gl["Credit"].sum()
        diff = abs(total_dr - total_cr)
        checks.append({
            "Check": "Debits = Credits (all data)",
            "Status": "PASS" if diff < 0.01 else "FAIL",
            "Detail": f"DR: ${total_dr:,.2f}  |  CR: ${total_cr:,.2f}  |  Diff: ${diff:,.2f}",
        })

    # Check 2: Account names match COA
    if coa:
        gl_accounts = set(gl["Account"].unique())
        active = coa.get("active_accounts", set())
        all_known = coa.get("account_names", set())
        unmatched = []
        for a in gl_accounts:
            a_clean = a.strip()
            if not a_clean or a_clean == "nan":
                continue
            # Fuzzy: check if GL account name is a substring of any COA name or vice versa
            found = False
            for known in all_known:
                if a_clean == known or a_clean in known or known in a_clean:
                    found = True
                    break
            if not found:
                # Try matching just the account number portion
                num_match = re.match(r"^(\d+)", a_clean)
                if num_match:
                    num = num_match.group(1)
                    for known in all_known:
                        if num in known:
                            found = True
                            break
            if not found:
                unmatched.append(a_clean)

        checks.append({
            "Check": "All GL accounts exist in COA",
            "Status": "PASS" if len(unmatched) == 0 else "WARN",
            "Detail": f"{len(unmatched)} unmatched account(s)" + (f": {', '.join(unmatched[:5])}" if unmatched else ""),
        })

        # Check 3: Active classes
        active_cls = coa.get("active_classes", set())
        if active_cls:
            checks.append({
                "Check": "Active classes loaded",
                "Status": "PASS",
                "Detail": f"{len(active_cls)} active class(es) available for validation",
            })
    else:
        checks.append({
            "Check": "COA validation",
            "Status": "SKIP",
            "Detail": "No Chart of Accounts uploaded — skipping account name validation",
        })

    # Check 4: No blank accounts
    blank_accts = gl["Account"].isna().sum() + (gl["Account"] == "").sum() + (gl["Account"] == "nan").sum()
    checks.append({
        "Check": "No blank account codes",
        "Status": "PASS" if blank_accts == 0 else "WARN",
        "Detail": f"{blank_accts} transaction(s) with blank account" if blank_accts else "All transactions have accounts",
    })

    # Check 5: No future dates
    if "Date" in gl.columns:
        future = gl[gl["Date"] > pd.Timestamp.now()].shape[0]
        checks.append({
            "Check": "No future-dated transactions",
            "Status": "PASS" if future == 0 else "WARN",
            "Detail": f"{future} future-dated transaction(s)" if future else "All dates are current or past",
        })

    # Display
    for ck in checks:
        status = ck["Status"]
        if status == "PASS":
            icon = pill("pass", "PASS")
        elif status == "FAIL":
            icon = pill("fail", "FAIL")
        else:
            icon = pill("warn", status)
        st.markdown(f"{icon} &nbsp; **{ck['Check']}** — {ck['Detail']}", unsafe_allow_html=True)

    return checks


def report_reconciliation(gl: pd.DataFrame, pdf_texts: list[str]):
    """Report 6 — Multi-Source Reconciliation Summary."""
    section("Multi-Source Reconciliation Summary")

    if not pdf_texts:
        st.info(
            "Upload bank statements or invoices (PDF) in the sidebar to enable "
            "three-way reconciliation. The system will attempt to match GL entries "
            "against extracted PDF line items."
        )

    # Bank-side: transactions with bank-related accounts
    bank_kw = r"(?i)(bank|cash|checking|savings|operating|deposit)"
    bank_gl = gl[gl["Account"].str.contains(bank_kw, na=False)].copy()

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(metric_card("GL Bank Transactions", str(len(bank_gl))), unsafe_allow_html=True)
    with c2:
        st.markdown(metric_card("PDF Documents", str(len(pdf_texts))), unsafe_allow_html=True)

    # Extract amounts from PDF text
    if pdf_texts:
        section("PDF-Extracted Line Items")
        all_amounts = []
        for idx, text in enumerate(pdf_texts):
            # Find dollar amounts in the text
            amounts = re.findall(r"\$?([\d,]+\.\d{2})", text)
            for a in amounts:
                val = float(a.replace(",", ""))
                if val > 0.01:
                    all_amounts.append({"Source": f"PDF #{idx+1}", "Amount": val})

        if all_amounts:
            pdf_df = pd.DataFrame(all_amounts)
            st.dataframe(pdf_df.head(50), use_container_width=True, hide_index=True)

            # Match against GL
            section("Reconciliation Matches")
            matches = []
            unmatched_pdf = []
            gl_amounts = bank_gl[["Date", "Account", "Memo", "Debit", "Credit"]].copy()

            for pa in all_amounts:
                val = pa["Amount"]
                # Try matching debit or credit
                match_dr = gl_amounts[(gl_amounts["Debit"] - val).abs() < 0.01]
                match_cr = gl_amounts[(gl_amounts["Credit"] - val).abs() < 0.01]
                if not match_dr.empty:
                    r = match_dr.iloc[0]
                    matches.append({
                        "PDF Amount": f"${val:,.2f}",
                        "GL Date": str(r["Date"])[:10],
                        "GL Account": r["Account"],
                        "GL Memo": str(r["Memo"])[:50],
                        "GL Amount (DR)": f"${r['Debit']:,.2f}",
                        "Status": "Matched",
                    })
                elif not match_cr.empty:
                    r = match_cr.iloc[0]
                    matches.append({
                        "PDF Amount": f"${val:,.2f}",
                        "GL Date": str(r["Date"])[:10],
                        "GL Account": r["Account"],
                        "GL Memo": str(r["Memo"])[:50],
                        "GL Amount (CR)": f"${r['Credit']:,.2f}",
                        "Status": "Matched",
                    })
                else:
                    unmatched_pdf.append({"PDF Amount": f"${val:,.2f}", "Status": "Unmatched"})

            if matches:
                st.dataframe(pd.DataFrame(matches), use_container_width=True, hide_index=True)
            if unmatched_pdf:
                st.markdown(f"**{len(unmatched_pdf)} PDF amount(s) unmatched** in the GL — these may represent "
                            "sales tax, shipping, or fees that banking rules missed.")
        else:
            st.warning("Could not extract numeric amounts from the uploaded PDF(s).")

    # Sales tax & shipping scan
    section("Sales Tax & Shipping Scan")
    tax_kw = r"(?i)(sales\s*tax|tax|shipping|freight|delivery|handling|surcharge)"
    tax_txns = gl[gl["Memo"].str.contains(tax_kw, na=False) | gl["Account"].str.contains(tax_kw, na=False)]
    if tax_txns.empty:
        st.info("No explicit sales tax or shipping entries found in the GL. "
                "These are common items that bank rules often miss — verify manually.")
    else:
        display = tax_txns[["Date", "Account", "Name", "Memo", "Debit", "Credit"]].copy()
        display["Date"] = display["Date"].astype(str).str[:10]
        for c in ["Debit", "Credit"]:
            display[c] = display[c].map(lambda v: f"${v:,.2f}" if v else "")
        st.dataframe(display, use_container_width=True, hide_index=True)


# ─────────────────────────────────────────────────────────────
# 4 · IIF EXPORT ENGINE
# ─────────────────────────────────────────────────────────────

def generate_iif(gl: pd.DataFrame, period=None) -> str:
    """Generate a QuickBooks-importable IIF file from GL transactions."""
    lines = []
    lines.append("!TRNS\tTRNSID\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO")
    lines.append("!SPL\tSPLID\tTRNSTYPE\tDATE\tACCNT\tNAME\tAMOUNT\tMEMO")
    lines.append("!ENDTRNS")

    if period is not None and "YearMonth" in gl.columns:
        data = gl[gl["YearMonth"] == period].copy()
    else:
        data = gl.copy()

    # Group by date + type to create journal entries
    if "Type" in data.columns:
        groups = data.groupby(["Date", "Type"])
    else:
        groups = data.groupby(["Date"])

    trns_id = 0
    for key, grp in groups:
        if len(grp) == 0:
            continue
        trns_id += 1
        first = grp.iloc[0]
        date_str = str(first["Date"])[:10] if pd.notna(first["Date"]) else ""
        trns_type = str(first.get("Type", "GENERAL JOURNAL")).strip()
        if trns_type in ("", "nan"):
            trns_type = "GENERAL JOURNAL"

        # First line = TRNS (the main entry)
        main_acct = str(first.get("Account", "")).replace("\t", " ")
        main_name = str(first.get("Name", "")).replace("\t", " ")
        if main_name == "nan":
            main_name = ""
        main_memo = str(first.get("Memo", "")).replace("\t", " ")
        if main_memo == "nan":
            main_memo = ""
        main_amount = first.get("Amount", 0)

        lines.append(f"TRNS\t{trns_id}\t{trns_type}\t{date_str}\t{main_acct}\t{main_name}\t{main_amount:.2f}\t{main_memo}")

        # Remaining lines = SPL
        for i in range(1, len(grp)):
            row = grp.iloc[i]
            acct = str(row.get("Account", "")).replace("\t", " ")
            name = str(row.get("Name", "")).replace("\t", " ")
            if name == "nan":
                name = ""
            memo = str(row.get("Memo", "")).replace("\t", " ")
            if memo == "nan":
                memo = ""
            amt = row.get("Amount", 0)
            lines.append(f"SPL\t{trns_id}\t{trns_type}\t{date_str}\t{acct}\t{name}\t{amt:.2f}\t{memo}")

        lines.append("ENDTRNS")

    return "\n".join(lines)


# ─────────────────────────────────────────────────────────────
# 5 · WORD DOCUMENT EXPORT ENGINE (python-docx)
# ─────────────────────────────────────────────────────────────

from docx import Document as DocxDocument
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml


def generate_docx_report(report_data: dict, output_path: str):
    """
    Generate an Apple-branded Word document from report data using python-docx.
    report_data keys:
      - title: str
      - subtitle: str
      - date: str
      - sections: list of {heading, paragraphs: list[str], table: {headers, rows}}
      - checks: list of {check, status, detail}  (for preflight)
    """
    doc = DocxDocument()

    # ── Page setup ──
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

    # ── Default font ──
    style = doc.styles["Normal"]
    style.font.name = "Helvetica"
    style.font.size = Pt(11)
    style.font.color.rgb = RGBColor(0x1D, 0x1D, 0x1F)

    # ── Header ──
    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    hp.clear()
    run1 = hp.add_run("Organization")
    run1.bold = True
    run1.font.size = Pt(8)
    run1.font.color.rgb = RGBColor(0x8E, 0x8E, 0x93)
    run2 = hp.add_run("  |  Month-End Close Agent")
    run2.font.size = Pt(8)
    run2.font.color.rgb = RGBColor(0x8E, 0x8E, 0x93)

    # ── Footer ──
    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    fp.clear()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = fp.add_run("Organization  |  Finance Department  |  Month-End Close Agent")
    fr.font.size = Pt(7)
    fr.font.color.rgb = RGBColor(0x8E, 0x8E, 0x93)

    # ── Title block ──
    doc.add_paragraph()  # spacer
    title_p = doc.add_paragraph()
    title_run = title_p.add_run(report_data.get("title", "Report"))
    title_run.bold = True
    title_run.font.size = Pt(24)
    title_run.font.color.rgb = RGBColor(0x1D, 0x1D, 0x1F)

    sub_p = doc.add_paragraph()
    sub_run = sub_p.add_run(report_data.get("subtitle", ""))
    sub_run.font.size = Pt(12)
    sub_run.font.color.rgb = RGBColor(0x6E, 0x6E, 0x73)

    date_str = report_data.get("date", datetime.date.today().strftime("%d %b %Y"))
    date_p = doc.add_paragraph()
    date_run = date_p.add_run(f"Generated: {date_str}")
    date_run.font.size = Pt(10)
    date_run.font.color.rgb = RGBColor(0x8E, 0x8E, 0x93)

    # Blue divider after title
    border_p = doc.add_paragraph()
    pPr = border_p._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}>'
                     f'<w:bottom w:val="single" w:sz="12" w:space="4" w:color="0071E3"/>'
                     f'</w:pBdr>')
    pPr.append(pBdr)

    doc.add_paragraph()  # spacer

    # ── Sections ──
    for sec_data in report_data.get("sections", []):
        heading_text = sec_data.get("heading", "")
        if heading_text:
            h_p = doc.add_paragraph()
            h_run = h_p.add_run(heading_text)
            h_run.bold = True
            h_run.font.size = Pt(14)
            h_run.font.color.rgb = RGBColor(0x00, 0x71, 0xE3)
            # Blue underline on heading
            hPr = h_p._p.get_or_add_pPr()
            hBdr = parse_xml(f'<w:pBdr {nsdecls("w")}>'
                             f'<w:bottom w:val="single" w:sz="8" w:space="4" w:color="0071E3"/>'
                             f'</w:pBdr>')
            hPr.append(hBdr)

        for para_text in sec_data.get("paragraphs", []):
            p = doc.add_paragraph()
            r = p.add_run(str(para_text))
            r.font.size = Pt(10.5)

        # Table
        table_data = sec_data.get("table")
        if table_data and table_data.get("headers") and table_data.get("rows"):
            headers = table_data["headers"]
            rows = table_data["rows"]
            n_cols = len(headers)
            tbl = doc.add_table(rows=1 + len(rows), cols=n_cols)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            tbl.autofit = True

            # Style header row
            for i, h in enumerate(headers):
                cell = tbl.rows[0].cells[i]
                cell.text = ""
                p = cell.paragraphs[0]
                r = p.add_run(str(h))
                r.bold = True
                r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                # Blue background
                shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="0071E3" w:val="clear"/>')
                cell._tc.get_or_add_tcPr().append(shading)

            # Data rows with alternating shading
            for ri, row in enumerate(rows):
                fill = "F5F5F7" if ri % 2 == 0 else "FFFFFF"
                for ci, cell_val in enumerate(row):
                    if ci >= n_cols:
                        break
                    cell = tbl.rows[ri + 1].cells[ci]
                    cell.text = ""
                    p = cell.paragraphs[0]
                    r = p.add_run(str(cell_val) if cell_val is not None else "")
                    r.font.size = Pt(9)
                    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill}" w:val="clear"/>')
                    cell._tc.get_or_add_tcPr().append(shading)

            # Light gray borders on all cells
            for row_obj in tbl.rows:
                for cell in row_obj.cells:
                    tcPr = cell._tc.get_or_add_tcPr()
                    borders = parse_xml(
                        f'<w:tcBorders {nsdecls("w")}>'
                        f'<w:top w:val="single" w:sz="2" w:space="0" w:color="D2D2D7"/>'
                        f'<w:bottom w:val="single" w:sz="2" w:space="0" w:color="D2D2D7"/>'
                        f'<w:left w:val="single" w:sz="2" w:space="0" w:color="D2D2D7"/>'
                        f'<w:right w:val="single" w:sz="2" w:space="0" w:color="D2D2D7"/>'
                        f'</w:tcBorders>'
                    )
                    tcPr.append(borders)

            doc.add_paragraph()  # spacer after table

    # ── Checks section (for preflight report) ──
    checks = report_data.get("checks", [])
    if checks:
        doc.add_paragraph()
        for ck in checks:
            icon = "\u2705" if ck["status"] == "PASS" else ("\u274C" if ck["status"] == "FAIL" else "\u26A0\uFE0F")
            ck_text = f'{icon} {ck["check"]} \u2014 {ck["detail"]}'
            p = doc.add_paragraph()
            r = p.add_run(ck_text)
            r.font.size = Pt(10)

    doc.save(output_path)


def build_flux_docx_data(gl: pd.DataFrame) -> dict:
    """Build report_data dict for the Flux Narrative report."""
    if "YearMonth" not in gl.columns:
        return None
    periods = sorted(gl["YearMonth"].dropna().unique())
    if len(periods) < 2:
        return None

    curr, prev = periods[-1], periods[-2]
    curr_df = gl[gl["YearMonth"] == curr]
    prev_df = gl[gl["YearMonth"] == prev]

    curr_totals = curr_df.groupby("Account")["Amount"].sum()
    prev_totals = prev_df.groupby("Account")["Amount"].sum()

    flux = pd.DataFrame({"Current": curr_totals, "Prior": prev_totals}).fillna(0)
    flux["Variance"] = flux["Current"] - flux["Prior"]
    flux["Var_%"] = np.where(flux["Prior"] != 0, (flux["Variance"] / flux["Prior"].abs()) * 100, np.nan)
    flux = flux.sort_values("Variance", key=abs, ascending=False)

    # Narrative paragraphs
    narratives = []
    for acct, row in flux.head(10).iterrows():
        direction = "increased" if row["Variance"] > 0 else "decreased"
        pct = f"{abs(row['Var_%']):.1f}%" if pd.notna(row["Var_%"]) else "N/A"
        acct_txns = curr_df[curr_df["Account"] == acct]
        vendors = [str(v) for v in acct_txns["Name"].dropna().unique()[:3] if str(v).strip() and str(v) != "nan"]
        memos = [str(m) for m in acct_txns["Memo"].dropna().unique()[:3] if str(m).strip() and str(m) != "nan"]
        vendor_str = ", ".join(vendors) if vendors else "various"
        memo_str = "; ".join(memos) if memos else "no memo detail"
        narratives.append(
            f"{acct} {direction} by ${abs(row['Variance']):,.2f} ({pct}), "
            f"from ${row['Prior']:,.2f} to ${row['Current']:,.2f}. "
            f"Key vendors: {vendor_str}. Memos: {memo_str}."
        )

    # Table
    table_rows = []
    for acct, row in flux.head(30).iterrows():
        table_rows.append([
            str(acct)[:40],
            f"${row['Current']:,.2f}",
            f"${row['Prior']:,.2f}",
            f"${row['Variance']:,.2f}",
            f"{row['Var_%']:.1f}%" if pd.notna(row["Var_%"]) else "N/A",
        ])

    return {
        "title": "Flux (Variance) Narrative Report",
        "subtitle": f"Organization  |  {prev} vs {curr}",
        "date": datetime.date.today().strftime("%d %b %Y"),
        "sections": [
            {"heading": "Variance Narrative", "paragraphs": narratives},
            {"heading": "Variance Detail",
             "paragraphs": [],
             "table": {
                 "headers": ["Account", "Current", "Prior", "Variance", "Var %"],
                 "rows": table_rows
             }},
        ]
    }


def build_vendor_gap_docx_data(gl: pd.DataFrame) -> dict:
    """Build report_data dict for Missing Bill report."""
    if "Name" not in gl.columns or "YearMonth" not in gl.columns:
        return None
    vendors = gl[gl["Name"].str.strip() != "nan"].copy()
    if vendors.empty:
        return None
    periods = sorted(vendors["YearMonth"].dropna().unique())
    if len(periods) < 3:
        return None

    curr = periods[-1]
    history = vendors[vendors["YearMonth"] != curr]
    curr_vendors = set(vendors[vendors["YearMonth"] == curr]["Name"].unique())
    total_hist = len(periods) - 1
    vendor_periods = history.groupby("Name")["YearMonth"].nunique()
    recurring = vendor_periods[vendor_periods >= max(2, total_hist * 0.5)]
    missing = [v for v in recurring.index if v not in curr_vendors and str(v).strip() and str(v) != "nan"]

    rows = []
    for v in missing:
        v_hist = history[history["Name"] == v]
        avg_amt = v_hist["Credit"].mean()
        last_date = v_hist["Date"].max()
        freq = v_hist["YearMonth"].nunique()
        typical_acct = v_hist["Account"].mode().iloc[0] if not v_hist["Account"].mode().empty else "Unknown"
        rows.append([
            str(v),
            f"{freq}/{total_hist} months",
            f"${avg_amt:,.2f}" if pd.notna(avg_amt) else "N/A",
            str(last_date.date()) if pd.notna(last_date) else "N/A",
            str(typical_acct)[:40],
            f"${avg_amt:,.2f}" if pd.notna(avg_amt) else "N/A",
        ])

    return {
        "title": "Recurring Vendor Gap Analysis",
        "subtitle": "Organization  |  Missing Bill Report",
        "date": datetime.date.today().strftime("%d %b %Y"),
        "sections": [
            {"heading": "Summary",
             "paragraphs": [
                 f"{len(recurring)} recurring vendors identified across {total_hist} months of history.",
                 f"{len(missing)} vendor(s) are missing from the current period ({curr}) and may require accrual entries.",
             ]},
            {"heading": "Suggested Accruals",
             "paragraphs": [],
             "table": {
                 "headers": ["Vendor", "Frequency", "Avg Amount", "Last Seen", "Typical Account", "Suggested Accrual"],
                 "rows": rows
             }},
        ]
    }


def build_suspense_docx_data(gl: pd.DataFrame, coa: dict | None) -> dict:
    """Build report_data for Suspense & Misc Reclass."""
    patterns = r"(?i)(suspense|misc|other|unclass|uncategoriz|clearing|unknown|unallocat)"
    mask = gl["Account"].str.contains(patterns, na=False) | gl["Memo"].str.contains(patterns, na=False)
    if "Split" in gl.columns:
        mask = mask | gl["Split"].str.contains(patterns, na=False)
    flagged = gl[mask].copy()

    active_accts = sorted(coa.get("active_accounts", set())) if coa else []
    rows = []
    for _, txn in flagged.iterrows():
        memo = str(txn.get("Memo", "")).lower()
        name = str(txn.get("Name", "")).lower()
        search_text = memo + " " + name
        best_match, best_score = "", 0
        for acct in active_accts:
            keywords = [k for k in re.split(r"[:\s\-&/]+", acct.lower()) if len(k) > 2]
            score = sum(1 for k in keywords if k in search_text)
            if score > best_score:
                best_score = score
                best_match = acct
        rows.append([
            str(txn.get("Date", ""))[:10],
            str(txn.get("Account", ""))[:35],
            str(txn.get("Name", ""))[:25],
            str(txn.get("Memo", ""))[:40],
            f"${txn['Amount']:,.2f}",
            best_match[:40] if best_score >= 1 else "Manual review",
            "High" if best_score >= 3 else ("Medium" if best_score >= 1 else "Low"),
        ])

    return {
        "title": "Suspense & Misc Resolution Worksheet",
        "subtitle": "Organization  |  Reclassification Recommendations",
        "date": datetime.date.today().strftime("%d %b %Y"),
        "sections": [
            {"heading": "Summary",
             "paragraphs": [
                 f"{len(flagged)} transaction(s) flagged in suspense, misc, clearing, or unclassified accounts.",
                 f"Total amount at risk: ${flagged['Amount'].abs().sum():,.2f}.",
             ]},
            {"heading": "Resolution Worksheet",
             "paragraphs": [],
             "table": {
                 "headers": ["Date", "Current Account", "Name", "Memo", "Amount", "Suggested Reclass", "Confidence"],
                 "rows": rows[:50]
             }},
        ]
    }


def build_materiality_docx_data(gl: pd.DataFrame, threshold: float) -> dict:
    """Build report_data for Materiality & Risk report."""
    large = gl[gl["Amount"].abs() >= threshold].copy()
    rows = []
    for _, txn in large.iterrows():
        risk = "Low"
        reasons = []
        acct = str(txn.get("Account", "")).lower()
        memo = str(txn.get("Memo", "")).lower()
        if re.search(r"(suspense|misc|other|clearing|unknown)", acct):
            risk = "High"
            reasons.append("suspense/misc account")
        if memo in ("", "nan", "none"):
            risk = "High" if risk == "High" else "Medium"
            reasons.append("no memo")
        if abs(txn["Amount"]) >= threshold * 5:
            risk = "High"
            reasons.append("exceeds 5x threshold")
        rows.append([
            str(txn.get("Date", ""))[:10],
            str(txn.get("Account", ""))[:35],
            str(txn.get("Name", ""))[:25],
            f"${txn['Amount']:,.2f}",
            risk,
            "; ".join(reasons) if reasons else "Material amount",
        ])

    return {
        "title": "Materiality & Risk Threshold Report",
        "subtitle": f"Organization  |  Threshold: ${threshold:,.0f}",
        "date": datetime.date.today().strftime("%d %b %Y"),
        "sections": [
            {"heading": "Summary",
             "paragraphs": [
                 f"Materiality threshold set at ${threshold:,.0f}.",
                 f"{len(rows)} transaction(s) exceed the threshold and have been flagged for review.",
                 f"High risk items: {sum(1 for r in rows if r[4] == 'High')}. Medium: {sum(1 for r in rows if r[4] == 'Medium')}.",
             ]},
            {"heading": "Flagged Transactions",
             "paragraphs": [],
             "table": {
                 "headers": ["Date", "Account", "Name", "Amount", "Risk", "Reason"],
                 "rows": rows[:50]
             }},
        ]
    }


def build_preflight_docx_data(gl: pd.DataFrame, coa: dict | None) -> dict:
    """Build report_data for IIF Pre-Flight."""
    checks = []
    if "YearMonth" in gl.columns:
        for period in sorted(gl["YearMonth"].dropna().unique()):
            p_df = gl[gl["YearMonth"] == period]
            total_dr, total_cr = p_df["Debit"].sum(), p_df["Credit"].sum()
            diff = abs(total_dr - total_cr)
            checks.append({
                "check": f"Debits = Credits ({period})",
                "status": "PASS" if diff < 0.01 else "FAIL",
                "detail": f"DR: ${total_dr:,.2f} | CR: ${total_cr:,.2f} | Diff: ${diff:,.2f}",
            })

    blank_accts = gl["Account"].isna().sum() + (gl["Account"] == "").sum() + (gl["Account"] == "nan").sum()
    checks.append({
        "check": "No blank account codes",
        "status": "PASS" if blank_accts == 0 else "WARN",
        "detail": f"{blank_accts} blank" if blank_accts else "All accounts populated",
    })

    if "Date" in gl.columns:
        future = gl[gl["Date"] > pd.Timestamp.now()].shape[0]
        checks.append({
            "check": "No future-dated transactions",
            "status": "PASS" if future == 0 else "WARN",
            "detail": f"{future} future-dated" if future else "All dates current",
        })

    return {
        "title": "IIF Import Pre-Flight Validation",
        "subtitle": "Organization  |  Technical Validation",
        "date": datetime.date.today().strftime("%d %b %Y"),
        "sections": [
            {"heading": "Validation Results",
             "paragraphs": [f"{sum(1 for c in checks if c['status']=='PASS')}/{len(checks)} checks passed."]}
        ],
        "checks": checks,
    }


def export_all_reports_docx(gl, coa, threshold, pdf_texts):
    """Generate all 6 reports as Word documents and return as a zip buffer."""
    import zipfile

    reports = {
        "01_Flux_Narrative": build_flux_docx_data(gl),
        "02_Missing_Bill_Analysis": build_vendor_gap_docx_data(gl),
        "03_Suspense_Resolution": build_suspense_docx_data(gl, coa),
        "04_Materiality_Risk": build_materiality_docx_data(gl, threshold),
        "05_IIF_PreFlight": build_preflight_docx_data(gl, coa),
    }

    # Report 6 - Reconciliation summary (simpler, text-based)
    bank_kw = r"(?i)(bank|cash|checking|savings|operating|deposit)"
    bank_gl = gl[gl["Account"].str.contains(bank_kw, na=False)]
    tax_kw = r"(?i)(sales\s*tax|tax|shipping|freight|delivery|handling)"
    tax_txns = gl[gl["Memo"].str.contains(tax_kw, na=False) | gl["Account"].str.contains(tax_kw, na=False)]

    recon_rows = []
    for _, txn in tax_txns.head(30).iterrows():
        recon_rows.append([
            str(txn.get("Date", ""))[:10],
            str(txn.get("Account", ""))[:35],
            str(txn.get("Memo", ""))[:40],
            f"${txn.get('Debit', 0):,.2f}" if txn.get("Debit", 0) else "",
            f"${txn.get('Credit', 0):,.2f}" if txn.get("Credit", 0) else "",
        ])

    reports["06_Reconciliation_Summary"] = {
        "title": "Multi-Source Reconciliation Summary",
        "subtitle": "Organization  |  Three-Way Reconciliation",
        "date": datetime.date.today().strftime("%d %b %Y"),
        "sections": [
            {"heading": "Overview",
             "paragraphs": [
                 f"GL bank transactions identified: {len(bank_gl):,}.",
                 f"PDF documents uploaded for reconciliation: {len(pdf_texts)}.",
                 f"Sales tax and shipping entries found: {len(tax_txns)}.",
             ]},
            {"heading": "Tax & Shipping Line Items",
             "paragraphs": [],
             "table": {
                 "headers": ["Date", "Account", "Memo", "Debit", "Credit"],
                 "rows": recon_rows
             }},
        ]
    }

    with tempfile.TemporaryDirectory() as tmpdir:
        paths = {}
        for name, data in reports.items():
            if data is None:
                continue
            path = os.path.join(tmpdir, f"{name}.docx")
            try:
                generate_docx_report(data, path)
                paths[name] = path
            except Exception as e:
                st.warning(f"Could not generate {name}: {e}")

        # Bundle into zip
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, path in paths.items():
                zf.write(path, f"Month-End Reports/{name}.docx")

        zip_buffer.seek(0)
        return zip_buffer, list(paths.keys())


# ─────────────────────────────────────────────────────────────
# 6 · MAIN APPLICATION
# ─────────────────────────────────────────────────────────────

def main():
    # Header
    st.markdown("""
    <div class="apple-header">
        <h1>Month-End Close Agent</h1>
        <p>Organization &nbsp;|&nbsp; Finance Department &nbsp;|&nbsp; Diagnostic Reporting Suite</p>
    </div>
    """, unsafe_allow_html=True)

    # ── Sidebar ──
    with st.sidebar:
        st.markdown("### Data Inputs")
        st.markdown("Upload your accounting source files to generate diagnostic reports.")

        gl_file = st.file_uploader(
            "General Ledger (GL)",
            type=["csv", "xlsx", "xls"],
            help="Transaction history export from QuickBooks (CSV or Excel).",
        )
        coa_file = st.file_uploader(
            "Chart of Accounts & Classes",
            type=["iif", "csv"],
            help="IIF export or CSV with account names and active classes.",
        )
        pdf_files = st.file_uploader(
            "Invoices / Bank Statements (optional)",
            type=["pdf"],
            accept_multiple_files=True,
            help="Upload PDFs for multi-source reconciliation.",
        )

        st.markdown("---")
        st.markdown("### Controls")
        threshold = st.slider(
            "Materiality Threshold ($)",
            min_value=100,
            max_value=50_000,
            value=1_000,
            step=100,
            help="Transactions exceeding this amount are flagged for audit risk review.",
        )

        st.markdown("---")
        st.markdown("### Export")
        export_iif = st.checkbox("Enable IIF Export", value=False)
        export_docx = st.checkbox("Export Reports as Word (.docx)", value=False,
                                  help="Generate Apple-branded Word documents for all 6 reports.")

    # ── Parse data ──
    gl = None
    coa = None
    pdf_texts = []

    if gl_file:
        with st.spinner("Parsing General Ledger..."):
            gl = parse_gl(gl_file)
        st.sidebar.success(f"GL loaded — {len(gl):,} transactions")

    if coa_file:
        with st.spinner("Parsing Chart of Accounts..."):
            if coa_file.name.lower().endswith(".iif"):
                coa = parse_iif(coa_file)
            else:
                coa = parse_coa_csv(coa_file)
        n_accts = len(coa.get("active_accounts", []))
        n_cls = len(coa.get("active_classes", []))
        st.sidebar.success(f"COA loaded — {n_accts} accounts, {n_cls} classes")

    if pdf_files:
        for pf in pdf_files:
            txt = extract_pdf_text(pf)
            if txt.strip():
                pdf_texts.append(txt)
        st.sidebar.success(f"{len(pdf_texts)} PDF(s) extracted")

    # ── Guard ──
    if gl is None:
        st.markdown("""
        <div style="text-align:center; padding:4rem 2rem;">
            <p style="font-size:1.1rem; color:#6E6E73;">
                Upload a <strong>General Ledger</strong> file in the sidebar to get started.
            </p>
            <p style="font-size:0.9rem; color:#8E8E93;">
                Supported formats: CSV, XLSX (QuickBooks Desktop export).
            </p>
        </div>
        """, unsafe_allow_html=True)
        return

    # ── Tabs ──
    tabs = st.tabs([
        "1 · Flux Narrative",
        "2 · Missing Bills",
        "3 · Suspense Reclass",
        "4 · Materiality",
        "5 · IIF Pre-Flight",
        "6 · Reconciliation",
    ])

    with tabs[0]:
        report_flux(gl)

    with tabs[1]:
        report_vendor_gap(gl)

    with tabs[2]:
        report_suspense(gl, coa)

    with tabs[3]:
        report_materiality(gl, threshold)

    with tabs[4]:
        checks = report_iif_preflight(gl, coa)

    with tabs[5]:
        report_reconciliation(gl, pdf_texts)

    # ── Word Document Export ──
    if export_docx:
        st.markdown("---")
        section("Word Document Export")
        if st.button("Generate All Reports as .docx", type="primary"):
            with st.spinner("Generating Apple-branded Word documents..."):
                try:
                    zip_buf, generated = export_all_reports_docx(gl, coa, threshold, pdf_texts)
                    st.success(f"Generated {len(generated)} report(s).")
                    st.download_button(
                        label="Download All Reports (.zip)",
                        data=zip_buf,
                        file_name=f"MonthEnd_Reports_{datetime.date.today().isoformat()}.zip",
                        mime="application/zip",
                    )
                except Exception as e:
                    st.error(f"Export failed: {e}")

    # ── IIF Export ──
    if export_iif:
        st.markdown("---")
        section("IIF Export")
        iif_content = generate_iif(gl)
        st.download_button(
            label="Download IIF File",
            data=iif_content,
            file_name=f"month_end_adjustments_{datetime.date.today().isoformat()}.iif",
            mime="text/plain",
        )
        with st.expander("Preview IIF Output"):
            st.code(iif_content[:3000], language="text")

    # Footer
    st.markdown(f"""
    <div class="apple-footer">
        Organization &nbsp;|&nbsp; Finance Department &nbsp;|&nbsp; Month-End Close Agent
        &nbsp;|&nbsp; Generated {datetime.date.today().strftime('%d %b %Y')}
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
