# Excel Data Consolidator V2

Streamlit application for consolidating multiple Excel files while preserving Excel formatting.

## What changed in V2

V1 used pandas to rebuild the output workbook. That preserved cell values but recreated the workbook, so original colors and formatting were lost.

V2 uses `openpyxl` and treats the **first uploaded Excel workbook as the master template**.

For every unique data row, V2 preserves:

- Cell fill/background color
- Font
- Borders
- Alignment
- Number/date formats
- Protection settings
- Hyperlinks
- Comments
- Row height (optional)

The first uploaded workbook preserves template-level properties such as:

- Worksheet name
- Header formatting
- Column widths
- Frozen panes
- Existing workbook sheets
- Page/layout settings supported by openpyxl
- AutoFilter, with its range updated for the consolidated table

## Duplicate handling

Two modes are available:

1. **Entire row** — all cell values must match.
2. **Selected columns** — choose one or more columns that determine whether a row is a duplicate.

The first occurrence is kept.

## Important assumptions

- All uploaded files must have the same columns and the same column order.
- Data begins immediately below the selected header row.
- V2 supports `.xlsx` and `.xlsm` uploads. Legacy `.xls` files are not supported because preserving modern Excel formatting reliably requires the XLSX/XLSM format.
- The first uploaded workbook controls workbook-level/template formatting.
- Unique rows contributed by later workbooks retain their own cell-level formatting.
- Complex features below the table, such as footers, merged cells inside the data area, external objects, advanced Excel tables, and macros may need special handling depending on the source workbook.

## Files

- `app.py`
- `requirements.txt`
- `README.md`
- `.gitignore`

## Streamlit Community Cloud deployment

1. Create or open your GitHub repository.
2. Replace the previous V1 files with the V2 files.
3. Commit and push.
4. Streamlit Cloud should automatically redeploy.
5. If necessary, reboot the app from the Streamlit dashboard.

Main file:

```text
app.py
```

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```
