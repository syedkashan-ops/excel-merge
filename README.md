# Excel Data Consolidator — Streamlit

## Features
- Upload multiple XLSX/XLS files.
- Consolidate files with the same column structure.
- Remove blank rows and duplicate rows.
- Optional duplicate detection using selected columns.
- Validate column structures before consolidation.
- Preview the consolidated data.
- Download one XLSX workbook containing `Consolidated Data` and `Summary` sheets.

## Deploy
Upload `app.py` and `requirements.txt` to GitHub, then deploy `app.py` with Streamlit Community Cloud.

## Local run
```bash
pip install -r requirements.txt
streamlit run app.py
```
