DAILY SALES MONITORING E-COMMERCE
===================================

Files:
1. app.py
2. eCommerce - Daily Sales Monitoring Dashboard.xlsx
3. requirements.txt

Run locally:
    pip install -r requirements.txt
    streamlit run app.py

Expected Excel sheets:
- DB Penjualan Produk 2026
- Monthly Sales Target

Expected sales columns:
- Tanggal
- Platform
- Product
- Sales Quantity
- Sales Value

Expected target columns:
- Year
- Month No
- Month
- Target Sales Value

Dashboard logic:
- Date = "as-of" date. MTD runs from the first day of that month through the selected date.
- Growth compares MTD against the same elapsed-day period in the previous month.
- Daily Sales Avg = MTD Sales Value / elapsed calendar days.
- Sales Target Achievement = MTD Sales Value / full monthly target.
- Target achievement is shown only when Platform = All Platform and Product = All Product because the database has no target breakdown by platform/product.
- TikTok/Tokopedia-like platform names are grouped as TikTok-Tokped.
- Shopee is grouped separately.
- Blibli, Lazada, Reseller, and other platforms are grouped as Others.
- Quick Insight calculates required daily sales = remaining gap / remaining calendar days.
