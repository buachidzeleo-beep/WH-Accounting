# Warehouse Stock App

Daily stock accounting for a 1C warehouse movement report ("ნომენკლატურის მოძრაობა საწყობზე").
A Streamlit app posts each day's movements — classified by the warehouse specialist with a
one-letter key (მ/დ/გ/ჩ) above each document column — into an accounting workbook
(balances + append-only ledger), matched on the 1C internal SKU code.

## Quick start

```bash
pip install -r requirements.txt
streamlit run app.py
```

Full usage instructions (Georgian): [README_SOP.md](README_SOP.md)

| File | Purpose |
|---|---|
| `app.py` | Streamlit UI |
| `engine.py` | Parsing / posting engine (UI-free, testable) |
| `სააღრიცხვო_ფაილი.xlsx` | Accounting file template (balances + empty ledger) |
| `README_SOP.md` | SOP for the warehouse specialist |
