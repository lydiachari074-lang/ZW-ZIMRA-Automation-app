"""
ZIMRA VAT Return Automation
A Streamlit app that turns a sales/purchases ledger into a ZIMRA-ready
VAT return: output tax, input tax (with denied/apportioned claims),
an audit trail, and a downloadable Excel return.
"""

import io
from datetime import datetime

import pandas as pd
import streamlit as st

VAT_RATE = 0.15  # ZIMRA standard VAT rate

REQUIRED_COLUMNS = ["date", "description", "type", "amount_excl_vat"]
OPTIONAL_COLUMNS = ["supply_type", "category", "mixed_use"]

# Keywords ZIMRA denies input tax on regardless of business purpose
ENTERTAINMENT_KEYWORDS = ["entertainment", "dinner", "restaurant", "hospitality", "gift"]
DENIED_VEHICLE_KEYWORDS = ["passenger motor vehicle", "passenger vehicle", "motor car", "sedan", "fortuner"]


# --------------------------------------------------------------------------
# Core VAT engine
# --------------------------------------------------------------------------

def load_and_validate(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Clean column names, check required columns, fill in optional ones."""
    errors = []
    df = df.copy()
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        errors.append(f"Missing required column(s): {', '.join(missing)}")
        return df, errors

    for col in OPTIONAL_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df["type"] = df["type"].astype(str).str.strip().str.lower()
    df["supply_type"] = df["supply_type"].fillna("").astype(str).str.strip().str.lower()
    df["category"] = df["category"].fillna("").astype(str).str.strip().str.lower()
    df["mixed_use"] = df["mixed_use"].fillna("").astype(str).str.strip().str.lower()
    df["description"] = df["description"].fillna("").astype(str)

    bad_type = ~df["type"].isin(["sale", "purchase"])
    if bad_type.any():
        errors.append(f"{bad_type.sum()} row(s) have a 'type' that is not 'sale' or 'purchase' and were dropped.")
        df = df[~bad_type]

    df["amount_excl_vat"] = pd.to_numeric(df["amount_excl_vat"], errors="coerce")
    bad_amount = df["amount_excl_vat"].isna()
    if bad_amount.any():
        errors.append(f"{bad_amount.sum()} row(s) have a non-numeric amount and were dropped.")
        df = df[~bad_amount]

    df["date"] = pd.to_datetime(df["date"], errors="coerce")

    return df.reset_index(drop=True), errors


def _is_denied_purchase(description: str, category: str) -> bool:
    text = f"{description} {category}".lower()
    if any(k in text for k in ENTERTAINMENT_KEYWORDS):
        return True
    if "capital goods" in category and any(k in text for k in DENIED_VEHICLE_KEYWORDS):
        return True
    return False


def compute_vat(df: pd.DataFrame, business_use_pct: float) -> pd.DataFrame:
    """Return the ledger with per-row VAT treatment and amounts."""
    rows = []
    for _, r in df.iterrows():
        supply_type = r["supply_type"]
        is_sale = r["type"] == "sale"
        amount = r["amount_excl_vat"]

        if supply_type == "exempt":
            rate = 0.0
            treatment = "Exempt supply/expense — no VAT"
            claim_pct = 0.0
        elif supply_type == "zero_rated":
            rate = 0.0
            treatment = "Zero-rated (0%)"
            claim_pct = 1.0
        else:
            rate = VAT_RATE
            treatment = "Standard-rated (15%)"
            claim_pct = 1.0

        denied_reason = ""
        if not is_sale:
            if _is_denied_purchase(r["description"], r["category"]):
                claim_pct = 0.0
                denied_reason = "Input tax denied (entertainment or passenger motor vehicle)"
                treatment = "Input tax denied"
            elif r["mixed_use"] == "yes":
                claim_pct = business_use_pct
                denied_reason = f"Mixed use — {business_use_pct:.0%} apportioned to business"
                treatment = f"Apportioned ({business_use_pct:.0%})"

        vat_amount = amount * rate
        claimable_or_output = vat_amount * claim_pct if not is_sale else vat_amount

        rows.append({
            **r.to_dict(),
            "vat_treatment": treatment,
            "vat_rate": rate,
            "vat_amount_full": round(vat_amount, 2),
            "claim_pct": claim_pct,
            "output_vat": round(vat_amount, 2) if is_sale else 0.0,
            "input_vat_claimable": round(claimable_or_output, 2) if not is_sale else 0.0,
            "input_vat_denied": round(vat_amount - claimable_or_output, 2) if not is_sale else 0.0,
            "notes": denied_reason,
        })

    return pd.DataFrame(rows)


def summarize(df: pd.DataFrame) -> dict:
    sales = df[df["type"] == "sale"]
    purchases = df[df["type"] == "purchase"]

    output_vat = sales["output_vat"].sum()
    input_vat = purchases["input_vat_claimable"].sum()
    input_vat_denied = purchases["input_vat_denied"].sum()

    return {
        "total_sales_excl_vat": sales["amount_excl_vat"].sum(),
        "standard_rated_sales": sales[sales["supply_type"] == ""]["amount_excl_vat"].sum(),
        "zero_rated_sales": sales[sales["supply_type"] == "zero_rated"]["amount_excl_vat"].sum(),
        "exempt_sales": sales[sales["supply_type"] == "exempt"]["amount_excl_vat"].sum(),
        "total_purchases_excl_vat": purchases["amount_excl_vat"].sum(),
        "output_vat": output_vat,
        "input_vat_claimable": input_vat,
        "input_vat_denied": input_vat_denied,
        "vat_payable": output_vat - input_vat,
        "transaction_count": len(df),
    }


def build_excel(df: pd.DataFrame, summary: dict, period_label: str) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        summary_df = pd.DataFrame([
            ["VAT Return Period", period_label],
            ["Generated", datetime.now().strftime("%Y-%m-%d %H:%M")],
            ["", ""],
            ["Total Sales (excl. VAT)", summary["total_sales_excl_vat"]],
            ["  Standard-rated sales", summary["standard_rated_sales"]],
            ["  Zero-rated sales", summary["zero_rated_sales"]],
            ["  Exempt sales", summary["exempt_sales"]],
            ["Total Purchases (excl. VAT)", summary["total_purchases_excl_vat"]],
            ["", ""],
            ["Output VAT (15% on standard-rated sales)", summary["output_vat"]],
            ["Input VAT claimable", summary["input_vat_claimable"]],
            ["Input VAT denied (entertainment/vehicles/exempt)", summary["input_vat_denied"]],
            ["NET VAT PAYABLE / (REFUNDABLE)", summary["vat_payable"]],
        ], columns=["Item", "Amount (USD)"])
        summary_df.to_excel(writer, sheet_name="VAT Return Summary", index=False)

        df[df["type"] == "sale"].to_excel(writer, sheet_name="Sales", index=False)
        df[df["type"] == "purchase"].to_excel(writer, sheet_name="Purchases", index=False)
        df.to_excel(writer, sheet_name="Audit Trail", index=False)

    buf.seek(0)
    return buf.getvalue()


# --------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------

st.set_page_config(page_title="ZIMRA VAT Return Automation", page_icon="🇿🇼", layout="wide")
st.title("🇿🇼 ZIMRA VAT Return Automation")
st.caption("Upload a sales & purchases ledger and get an audit-ready VAT return in seconds.")

with st.sidebar:
    st.header("Settings")
    business_use_pct = st.slider(
        "Default business-use % for mixed-use expenses",
        min_value=0, max_value=100, value=50, step=5,
        help="Applied to purchases flagged mixed_use=yes (e.g. shared electricity, phone bills).",
    ) / 100
    period_label = st.text_input("VAT period (for the return header)", value=datetime.now().strftime("%B %Y"))
    st.divider()
    st.markdown(
        "**Denied input tax rules applied automatically:**\n"
        "- Entertainment expenses\n"
        "- Passenger motor vehicles (capital goods)\n"
        "- Purchases marked `exempt`\n"
        "- Mixed-use expenses are apportioned at the % above"
    )

uploaded = st.file_uploader("Upload transactions (CSV or Excel)", type=["csv", "xlsx", "xls"])

use_sample = False
if uploaded is None:
    use_sample = st.checkbox("No file? Use the bundled sample_transactions.csv", value=True)

raw_df = None
if uploaded is not None:
    if uploaded.name.endswith((".xlsx", ".xls")):
        raw_df = pd.read_excel(uploaded)
    else:
        raw_df = pd.read_csv(uploaded)
elif use_sample:
    try:
        raw_df = pd.read_csv("sample_transactions.csv")
    except FileNotFoundError:
        st.warning("sample_transactions.csv not found next to app.py.")

if raw_df is None:
    st.info("Upload a file to get started, or tick the sample-data box above.")
    with st.expander("Required file format"):
        st.markdown(
            "Columns: `date, description, type, amount_excl_vat, supply_type, category, mixed_use`\n\n"
            "- `type`: `sale` or `purchase`\n"
            "- `supply_type`: blank (standard-rated), `zero_rated`, or `exempt`\n"
            "- `category`: free text, e.g. `stock`, `overheads`, `capital goods`, `rental`\n"
            "- `mixed_use`: `yes` if the expense is shared between business and private/exempt use"
        )
    st.stop()

df, errors = load_and_validate(raw_df)
for e in errors:
    st.warning(e)

if df.empty:
    st.error("No valid transactions to process.")
    st.stop()

result = compute_vat(df, business_use_pct)
summary = summarize(result)

st.subheader("1. Summary")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Total Sales (excl. VAT)", f"${summary['total_sales_excl_vat']:,.2f}")
c2.metric("Total Purchases (excl. VAT)", f"${summary['total_purchases_excl_vat']:,.2f}")
c3.metric("Transactions", summary["transaction_count"])
c4.metric("VAT Payable" if summary["vat_payable"] >= 0 else "VAT Refundable",
          f"${abs(summary['vat_payable']):,.2f}")

st.subheader("2. VAT Return")
v1, v2, v3 = st.columns(3)
v1.metric("Output VAT", f"${summary['output_vat']:,.2f}")
v2.metric("Input VAT Claimable", f"${summary['input_vat_claimable']:,.2f}")
v3.metric("Input VAT Denied", f"${summary['input_vat_denied']:,.2f}")

if summary["input_vat_denied"] > 0:
    st.caption(
        f"⚠️ ${summary['input_vat_denied']:,.2f} of input VAT was denied — see the Notes column "
        "in the audit trail below for why."
    )

st.subheader("3. Sales Breakdown")
st.dataframe(pd.DataFrame([
    {"Supply type": "Standard-rated (15%)", "Amount (excl. VAT)": summary["standard_rated_sales"]},
    {"Supply type": "Zero-rated (0%)", "Amount (excl. VAT)": summary["zero_rated_sales"]},
    {"Supply type": "Exempt", "Amount (excl. VAT)": summary["exempt_sales"]},
]), hide_index=True, use_container_width=True)

st.subheader("4. Audit Trail")
display_cols = ["date", "description", "type", "amount_excl_vat", "vat_treatment",
                 "output_vat", "input_vat_claimable", "input_vat_denied", "notes"]
st.dataframe(result[display_cols], hide_index=True, use_container_width=True)

st.subheader("5. Download VAT Return")
try:
    excel_bytes = build_excel(result, summary, period_label)
    st.download_button(
        "📥 Download VAT Return (Excel)",
        data=excel_bytes,
        file_name=f"ZIMRA_VAT_Return_{period_label.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
except ModuleNotFoundError:
    st.error(
        "The Excel export needs the `openpyxl` package, which isn't installed on this "
        "deployment. Add `openpyxl` to requirements.txt at the repo root, commit it, then "
        "reboot the app from **Manage app → Reboot app** on Streamlit Cloud so it reinstalls "
        "dependencies. As a fallback, you can still copy the tables above manually."
    )
    st.download_button(
        "📥 Download Audit Trail (CSV fallback)",
        data=result.to_csv(index=False).encode("utf-8"),
        file_name=f"ZIMRA_VAT_AuditTrail_{period_label.replace(' ', '_')}.csv",
        mime="text/csv",
    )
