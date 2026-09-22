# 🇿🇼 ZIMRA VAT Return Automation

A Streamlit app that turns a sales & purchases ledger into a ZIMRA-ready VAT
return: output tax, input tax (with denied/apportioned claims applied
automatically), a full audit trail, and a downloadable Excel VAT return.

## Features

- Upload your own CSV/Excel ledger, or try the bundled sample data
- Handles **standard-rated (15%)**, **zero-rated (0%)**, and **exempt** supplies
- Automatically **denies input VAT** on entertainment expenses and passenger
  motor vehicles, per ZIMRA rules
- **Apportions** mixed-use expenses (e.g. shared electricity) at an
  adjustable business-use percentage
- Full audit trail showing the VAT treatment and reasoning for every line
- One-click **Excel VAT return** download (Summary, Sales, Purchases, Audit Trail sheets)
- Validates uploaded files and flags bad rows instead of crashing

## How to Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually http://localhost:8501).

## File Format

Required columns: `date, description, type, amount_excl_vat`
Optional columns: `supply_type, category, mixed_use`

| Column | Values | Notes |
|---|---|---|
| `type` | `sale` or `purchase` | required |
| `supply_type` | blank / `zero_rated` / `exempt` | blank = standard-rated (15%) |
| `category` | free text (`stock`, `overheads`, `capital goods`, `rental`, ...) | used with description to detect denied claims |
| `mixed_use` | `yes` or blank | triggers apportionment at the sidebar % |

## VAT Logic Applied

- **Output VAT** = 15% × standard-rated sales (zero-rated and exempt sales are
  reported but carry no output VAT)
- **Input VAT** is claimable in full unless:
  - the purchase description/category matches **entertainment** keywords → denied
  - the purchase is a **passenger motor vehicle** under `capital goods` → denied
  - the purchase is marked `exempt` → no VAT was charged, nothing to claim
  - the purchase is `mixed_use` → only the business-use % is claimed
- **VAT Payable** = Output VAT − Input VAT Claimable

This mirrors the core ZIMRA VAT201 logic but is a simplification for
demonstration/bookkeeping support — always reconcile against ZIMRA's current
VAT Act and e-filing portal before submitting a return.

## Deploying

- **Streamlit Community Cloud**: push this repo to GitHub, then deploy at
  https://share.streamlit.io by pointing it at `app.py`.
- **GitHub**: see below for pushing this project to your own repository.

## Pushing to GitHub

```bash
cd vat-app
git init
git add .
git commit -m "ZIMRA VAT return automation app"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```
