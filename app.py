import io
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Excel Data Consolidator', page_icon='📊', layout='wide')
st.title('📊 Excel Data Consolidator')
st.caption('Upload multiple Excel files with the same structure, consolidate them, remove duplicates, and download one Excel file.')

with st.sidebar:
    st.header('Settings')
    duplicate_mode = st.radio('Duplicate handling', ['Entire row', 'Selected columns'])
    sheet_mode = st.radio('Excel sheet', ['First sheet', 'Select a sheet'])

files = st.file_uploader('Upload Excel files', type=['xlsx', 'xls'], accept_multiple_files=True)
if not files:
    st.info('Upload Excel files to begin.')
    st.stop()

selected_sheet = None
if sheet_mode == 'Select a sheet':
    names = set()
    for f in files:
        try:
            f.seek(0)
            names.update(pd.ExcelFile(f).sheet_names)
        except Exception as e:
            st.warning(f'{f.name}: {e}')
        finally:
            f.seek(0)
    if not names:
        st.error('No readable Excel sheets found.')
        st.stop()
    selected_sheet = st.selectbox('Select the sheet to consolidate', sorted(names))

dfs, stats, errors = [], [], []
for f in files:
    try:
        f.seek(0)
        sheet = 0 if sheet_mode == 'First sheet' else selected_sheet
        df = pd.read_excel(f, sheet_name=sheet)
        df.columns = [str(c).strip() for c in df.columns]
        before = len(df)
        df = df.dropna(how='all').copy()
        dfs.append((f.name, df))
        stats.append({'File': f.name, 'Rows read': before, 'Blank rows removed': before-len(df), 'Usable rows': len(df), 'Columns': len(df.columns)})
    except Exception as e:
        errors.append(f'{f.name}: {e}')

if errors:
    st.error('Some files could not be read:')
    for e in errors: st.write(f'- {e}')
if not dfs:
    st.error('No usable Excel files were loaded.')
    st.stop()

reference_columns = list(dfs[0][1].columns)
mismatches = []
for name, df in dfs[1:]:
    cols = list(df.columns)
    if cols != reference_columns:
        missing = [c for c in reference_columns if c not in cols]
        extra = [c for c in cols if c not in reference_columns]
        mismatches.append({'File': name, 'Missing columns': ', '.join(missing), 'Extra columns': ', '.join(extra), 'Same columns but different order': 'Yes' if not missing and not extra else 'No'})
if mismatches:
    st.warning('Column structures are not identical. Consolidation was stopped to prevent incorrect data alignment.')
    st.dataframe(pd.DataFrame(mismatches), use_container_width=True, hide_index=True)
    st.stop()

combined = pd.concat([df for _, df in dfs], ignore_index=True)
total_rows = len(combined)

if duplicate_mode == 'Entire row':
    mask = combined.duplicated(keep='first')
    consolidated = combined.loc[~mask].copy()
    duplicate_count = int(mask.sum())
else:
    selected_columns = st.multiselect('Select duplicate-identifying columns', reference_columns)
    if selected_columns:
        mask = combined.duplicated(subset=selected_columns, keep='first')
        consolidated = combined.loc[~mask].copy()
        duplicate_count = int(mask.sum())
    else:
        st.info('Select at least one column to enable duplicate removal by selected columns.')
        consolidated = combined.copy()
        duplicate_count = 0

c1,c2,c3,c4 = st.columns(4)
c1.metric('Files processed', len(dfs))
c2.metric('Rows imported', f'{total_rows:,}')
c3.metric('Duplicates removed', f'{duplicate_count:,}')
c4.metric('Final unique rows', f'{len(consolidated):,}')

with st.expander('File processing summary'):
    st.dataframe(pd.DataFrame(stats), use_container_width=True, hide_index=True)
with st.expander('Detected columns'):
    st.write(reference_columns)

st.subheader('Preview')
max_preview = min(100, max(5, len(consolidated)))
preview = st.slider('Rows to preview', 5, max_preview, min(20, max_preview))
st.dataframe(consolidated.head(preview), use_container_width=True, hide_index=True)

output = io.BytesIO()
filename = f"Consolidated_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    consolidated.to_excel(writer, index=False, sheet_name='Consolidated Data')
    pd.DataFrame({'Metric':['Files processed','Rows imported','Duplicates removed','Final unique rows','Generated'], 'Value':[len(dfs),total_rows,duplicate_count,len(consolidated),datetime.now().strftime('%Y-%m-%d %H:%M:%S')]}).to_excel(writer,index=False,sheet_name='Summary')
    ws = writer.book['Consolidated Data']
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]: cell.font = cell.font.copy(bold=True)
    for cells in ws.columns:
        letter=cells[0].column_letter
        max_len=max(len(str(c.value or '')) for c in cells[:200])
        ws.column_dimensions[letter].width=min(max(max_len+2,10),45)
    writer.book['Summary'].column_dimensions['A'].width=28
    writer.book['Summary'].column_dimensions['B'].width=28

st.download_button('⬇️ Download Consolidated Excel', output.getvalue(), filename, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', type='primary', use_container_width=True)
st.caption('Uploaded files are processed in memory and are not modified.')
