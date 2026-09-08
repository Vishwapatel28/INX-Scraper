# India INX Scraper

1. Open PowerShell in this folder.
2. Run:
   `python -m pip install -r requirements.txt`
3. Run:
   `streamlit run app.py`
4. Use the UI.
5. The final Excel is created in `output\` with:
   `India_INX_Members_DD-MM-YY.xlsx`

The workbook contains exactly four sheets:
Members, Directors, Designated, Others.

Chrome runs headless, so no visible Chrome popup is used.
"All Members" follows pagination automatically.
"Selected Pages" lets you test a page range.
