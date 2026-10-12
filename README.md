# Cutting WIP Control Tower

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

The sample workbook is included. You can also upload another HMES Cutting WIP Inquiry Excel export from the sidebar.

## Features
- Order, cut, bundle and issue KPI cards
- WIP funnel and customer/season analysis
- Sales-order priority board
- Quantity and balance exception detection
- Filtered Excel export

## Important
The source dataset has no transaction timestamps, so this app does not calculate hourly WIP aging. Priority thresholds are business assumptions and can be changed in `app.py`.

## Excel style compatibility
The application uses the Calamine reader first, allowing it to read HMES exports even when their Excel style metadata is malformed.

## Pandas compatibility
Priority-cell formatting supports both newer pandas versions (`Styler.map`) and older versions (`Styler.applymap`).
