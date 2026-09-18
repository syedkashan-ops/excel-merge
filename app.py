import io
from copy import copy
from datetime import datetime

import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="Excel Data Consolidator V2",
    page_icon="📊",
    layout="wide",
)

st.markdown(
    """
    <style>
    .main-title {
        font-size: 2.25rem;
        font-weight: 750;
        margin-bottom: 0.15rem;
    }
    .sub-title {
        color: #666;
        margin-bottom: 1.25rem;
    }
    .note {
        padding: .85rem 1rem;
        border: 1px solid rgba(128,128,128,.25);
        border-radius: 10px;
        margin-bottom: 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="main-title">📊 Excel Data Consolidator V2</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-title">Consolidate Excel files while preserving formatting, colors, borders, fonts, alignment and number formats.</div>',
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="note">
    <b>V2 formatting behavior:</b> The first uploaded workbook is used as the master template.
    Each unique data row keeps the cell formatting from the file it originally came from.
    </div>
    """,
    unsafe_allow_html=True,
)


# -----------------------------
# Helpers
# -----------------------------
def normalized_value(value):
    """Create a stable comparison value for duplicate detection."""
    if isinstance(value, datetime):
        return ("datetime", value.isoformat())
    try:
        # Dates can also arrive as date objects.
        if hasattr(value, "isoformat") and not isinstance(value, (str, bytes)):
            return (type(value).__name__, value.isoformat())
    except Exception:
        pass

    if isinstance(value, str):
        # Intentionally do not lowercase or strip internal content:
        # the goal is exact-value duplicate detection.
        return ("str", value)

    return (type(value).__name__, value)


def row_key(values, indexes=None):
    if indexes is None:
        chosen = values
    else:
        chosen = [values[i] for i in indexes]
    return tuple(normalized_value(v) for v in chosen)


def row_is_blank(values):
    return all(v is None or (isinstance(v, str) and v == "") for v in values)


def copy_cell_format(source_cell, target_cell):
    """Copy Excel cell style and useful metadata without sharing style objects."""
    if source_cell.has_style:
        target_cell._style = copy(source_cell._style)

    if source_cell.number_format:
        target_cell.number_format = source_cell.number_format

    target_cell.font = copy(source_cell.font)
    target_cell.fill = copy(source_cell.fill)
    target_cell.border = copy(source_cell.border)
    target_cell.alignment = copy(source_cell.alignment)
    target_cell.protection = copy(source_cell.protection)

    if source_cell.hyperlink:
        target_cell._hyperlink = copy(source_cell.hyperlink)

    if source_cell.comment:
        target_cell.comment = copy(source_cell.comment)


def read_uploaded_bytes(uploaded_file):
    uploaded_file.seek(0)
    return uploaded_file.read()


def load_uploaded_workbook(file_bytes):
    return load_workbook(
        io.BytesIO(file_bytes),
        data_only=False,
        keep_links=True,
    )


# -----------------------------
# Upload
# -----------------------------
uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
    help="V2 accepts XLSX/XLSM because formatting preservation requires the modern Excel workbook format.",
)

if not uploaded_files:
    st.info("Upload two or more .xlsx/.xlsm files to begin.")
    st.stop()

if len(uploaded_files) < 2:
    st.warning("You can continue with one file, but this app is designed for consolidating multiple files.")


# Keep bytes in memory because Streamlit UploadedFile streams can move position.
uploaded = []
load_errors = []

for f in uploaded_files:
    try:
        raw = read_uploaded_bytes(f)
        wb = load_uploaded_workbook(raw)
        uploaded.append(
            {
                "name": f.name,
                "bytes": raw,
                "sheet_names": wb.sheetnames,
            }
        )
        wb.close()
    except Exception as exc:
        load_errors.append(f"{f.name}: {exc}")

if load_errors:
    st.error("Some files could not be opened:")
    for item in load_errors:
        st.write(f"- {item}")

if not uploaded:
    st.stop()


# -----------------------------
# Settings
# -----------------------------
with st.sidebar:
    st.header("Consolidation Settings")

    first_sheets = uploaded[0]["sheet_names"]
    selected_sheet = st.selectbox(
        "Sheet to consolidate",
        first_sheets,
        index=0,
        help="The same sheet name must exist in every uploaded workbook.",
    )

    header_row = st.number_input(
        "Header row number",
        min_value=1,
        max_value=100,
        value=1,
        step=1,
        help="Data is assumed to start immediately below this row.",
    )

    duplicate_mode = st.radio(
        "Duplicate detection",
        ["Entire row", "Selected columns"],
        index=0,
    )

    preserve_source_row_height = st.checkbox(
        "Preserve source row heights",
        value=True,
    )


# Check selected sheet exists everywhere.
missing_sheet = [
    item["name"]
    for item in uploaded
    if selected_sheet not in item["sheet_names"]
]

if missing_sheet:
    st.error(
        f'The sheet "{selected_sheet}" is missing from: '
        + ", ".join(missing_sheet)
    )
    st.stop()


# -----------------------------
# Inspect headers
# -----------------------------
workbook_meta = []
header_lists = []
inspection_errors = []

for item in uploaded:
    try:
        wb = load_uploaded_workbook(item["bytes"])
        ws = wb[selected_sheet]

        headers = [
            ws.cell(int(header_row), col).value
            for col in range(1, ws.max_column + 1)
        ]

        # Remove trailing completely empty header columns.
        while headers and headers[-1] is None:
            headers.pop()

        header_lists.append(headers)
        workbook_meta.append(
            {
                "File": item["name"],
                "Sheet": selected_sheet,
                "Rows": ws.max_row,
                "Columns": len(headers),
            }
        )
        wb.close()

    except Exception as exc:
        inspection_errors.append(f"{item['name']}: {exc}")

if inspection_errors:
    st.error("Could not inspect one or more workbooks:")
    for item in inspection_errors:
        st.write(f"- {item}")
    st.stop()

if not header_lists or not header_lists[0]:
    st.error("No headers were found on the selected header row.")
    st.stop()

reference_headers = header_lists[0]


# Exact header structure validation.
structure_errors = []

for idx, headers in enumerate(header_lists[1:], start=1):
    if headers != reference_headers:
        ref_set = set(reference_headers)
        cur_set = set(headers)

        missing = [str(h) for h in reference_headers if h not in cur_set]
        extra = [str(h) for h in headers if h not in ref_set]

        structure_errors.append(
            {
                "File": uploaded[idx]["name"],
                "Missing columns": ", ".join(missing),
                "Extra columns": ", ".join(extra),
                "Different order": "Yes" if not missing and not extra else "N/A",
            }
        )

if structure_errors:
    st.error(
        "The uploaded files do not have an identical column structure. "
        "Consolidation has been stopped to avoid incorrect data placement."
    )
    st.dataframe(structure_errors, use_container_width=True, hide_index=True)
    st.stop()


# Duplicate column selection after headers are known.
duplicate_indexes = None

if duplicate_mode == "Selected columns":
    display_headers = [
        f"{i + 1}. {h if h is not None else '(blank header)'}"
        for i, h in enumerate(reference_headers)
    ]

    selected_display = st.multiselect(
        "Choose columns that define a duplicate",
        options=display_headers,
        default=[],
    )

    if not selected_display:
        st.info("Choose at least one duplicate-identifying column.")
        st.stop()

    duplicate_indexes = [
        display_headers.index(label)
        for label in selected_display
    ]


# -----------------------------
# Collect unique rows + styles
# -----------------------------
seen = set()
unique_rows = []
file_stats = []

total_nonblank_rows = 0
duplicate_count = 0

for item in uploaded:
    wb = load_uploaded_workbook(item["bytes"])
    ws = wb[selected_sheet]

    usable = 0
    duplicates_here = 0
    blanks_here = 0

    data_start = int(header_row) + 1
    last_col = len(reference_headers)

    for source_row_num in range(data_start, ws.max_row + 1):
        source_cells = [
            ws.cell(source_row_num, col)
            for col in range(1, last_col + 1)
        ]
        values = [cell.value for cell in source_cells]

        if row_is_blank(values):
            blanks_here += 1
            continue

        usable += 1
        total_nonblank_rows += 1

        key = row_key(values, duplicate_indexes)

        if key in seen:
            duplicate_count += 1
            duplicates_here += 1
            continue

        seen.add(key)

        unique_rows.append(
            {
                "source_file": item["name"],
                "source_row": source_row_num,
                "values": values,
                # Preserve style objects while workbook is still open.
                "styles": [copy(c._style) for c in source_cells],
                "fonts": [copy(c.font) for c in source_cells],
                "fills": [copy(c.fill) for c in source_cells],
                "borders": [copy(c.border) for c in source_cells],
                "alignments": [copy(c.alignment) for c in source_cells],
                "protections": [copy(c.protection) for c in source_cells],
                "number_formats": [c.number_format for c in source_cells],
                "hyperlinks": [copy(c.hyperlink) if c.hyperlink else None for c in source_cells],
                "comments": [copy(c.comment) if c.comment else None for c in source_cells],
                "row_height": ws.row_dimensions[source_row_num].height,
                "row_hidden": ws.row_dimensions[source_row_num].hidden,
            }
        )

    file_stats.append(
        {
            "File": item["name"],
            "Non-blank rows": usable,
            "Duplicates ignored": duplicates_here,
            "Blank rows ignored": blanks_here,
            "Unique rows contributed": usable - duplicates_here,
        }
    )
    wb.close()


# -----------------------------
# Build formatted output
# -----------------------------
template = uploaded[0]
output_wb = load_uploaded_workbook(template["bytes"])
output_ws = output_wb[selected_sheet]

data_start = int(header_row) + 1
last_col = len(reference_headers)

# Capture the template's original data-area style as a fallback.
fallback_styles = []
fallback_height = None

if output_ws.max_row >= data_start:
    fallback_height = output_ws.row_dimensions[data_start].height
    for col in range(1, last_col + 1):
        c = output_ws.cell(data_start, col)
        fallback_styles.append(
            {
                "style": copy(c._style),
                "font": copy(c.font),
                "fill": copy(c.fill),
                "border": copy(c.border),
                "alignment": copy(c.alignment),
                "protection": copy(c.protection),
                "number_format": c.number_format,
            }
        )

# Remove the template's old data rows.
# This produces a clean consolidated table while retaining everything above the header
# and workbook/sheet-level properties from the first file.
if output_ws.max_row >= data_start:
    output_ws.delete_rows(data_start, output_ws.max_row - data_start + 1)

# Recreate unique rows with their original source formatting.
for offset, row_data in enumerate(unique_rows):
    target_row_num = data_start + offset

    for idx, value in enumerate(row_data["values"], start=1):
        target = output_ws.cell(target_row_num, idx)
        target.value = value

        source_index = idx - 1
        target._style = copy(row_data["styles"][source_index])
        target.font = copy(row_data["fonts"][source_index])
        target.fill = copy(row_data["fills"][source_index])
        target.border = copy(row_data["borders"][source_index])
        target.alignment = copy(row_data["alignments"][source_index])
        target.protection = copy(row_data["protections"][source_index])
        target.number_format = row_data["number_formats"][source_index]

        if row_data["hyperlinks"][source_index]:
            target._hyperlink = copy(row_data["hyperlinks"][source_index])

        if row_data["comments"][source_index]:
            target.comment = copy(row_data["comments"][source_index])

    if preserve_source_row_height:
        output_ws.row_dimensions[target_row_num].height = row_data["row_height"]
        output_ws.row_dimensions[target_row_num].hidden = row_data["row_hidden"]
    elif fallback_height is not None:
        output_ws.row_dimensions[target_row_num].height = fallback_height


# Update AutoFilter range if the sheet already uses one.
if output_ws.auto_filter and output_ws.auto_filter.ref:
    end_row = max(int(header_row), data_start + len(unique_rows) - 1)
    output_ws.auto_filter.ref = (
        f"A{int(header_row)}:{get_column_letter(last_col)}{end_row}"
    )


# -----------------------------
# Summary sheet
# -----------------------------
summary_name = "Consolidation Summary"

if summary_name in output_wb.sheetnames:
    del output_wb[summary_name]

summary_ws = output_wb.create_sheet(summary_name)

summary_ws["A1"] = "Excel Data Consolidator V2"
summary_ws["A1"].font = copy(output_ws.cell(int(header_row), 1).font)
summary_ws["A3"] = "Files processed"
summary_ws["B3"] = len(uploaded)
summary_ws["A4"] = "Rows imported"
summary_ws["B4"] = total_nonblank_rows
summary_ws["A5"] = "Duplicates ignored"
summary_ws["B5"] = duplicate_count
summary_ws["A6"] = "Final unique rows"
summary_ws["B6"] = len(unique_rows)
summary_ws["A7"] = "Template workbook"
summary_ws["B7"] = template["name"]
summary_ws["A8"] = "Template sheet"
summary_ws["B8"] = selected_sheet
summary_ws["A9"] = "Header row"
summary_ws["B9"] = int(header_row)
summary_ws["A10"] = "Generated"
summary_ws["B10"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

summary_ws["A12"] = "Per-file processing"
summary_headers = [
    "File",
    "Non-blank rows",
    "Duplicates ignored",
    "Blank rows ignored",
    "Unique rows contributed",
]

for col, title in enumerate(summary_headers, start=1):
    summary_ws.cell(13, col).value = title
    summary_ws.cell(13, col).font = copy(output_ws.cell(int(header_row), min(col, last_col)).font)
    summary_ws.cell(13, col).fill = copy(output_ws.cell(int(header_row), min(col, last_col)).fill)
    summary_ws.cell(13, col).border = copy(output_ws.cell(int(header_row), min(col, last_col)).border)
    summary_ws.cell(13, col).alignment = copy(output_ws.cell(int(header_row), min(col, last_col)).alignment)

for r, record in enumerate(file_stats, start=14):
    for c, title in enumerate(summary_headers, start=1):
        summary_ws.cell(r, c).value = record[title]

summary_ws.freeze_panes = "A14"
summary_ws.column_dimensions["A"].width = 38
for col in range(2, 6):
    summary_ws.column_dimensions[get_column_letter(col)].width = 22


# -----------------------------
# Save to memory
# -----------------------------
out = io.BytesIO()
output_wb.save(out)
output_wb.close()
out.seek(0)

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_filename = f"Consolidated_Formatted_{timestamp}.xlsx"


# -----------------------------
# UI results
# -----------------------------
st.divider()

c1, c2, c3, c4 = st.columns(4)
c1.metric("Files processed", len(uploaded))
c2.metric("Rows imported", f"{total_nonblank_rows:,}")
c3.metric("Duplicates ignored", f"{duplicate_count:,}")
c4.metric("Final unique rows", f"{len(unique_rows):,}")

with st.expander("File processing details", expanded=False):
    st.dataframe(
        file_stats,
        use_container_width=True,
        hide_index=True,
    )

with st.expander("Detected columns", expanded=False):
    for i, h in enumerate(reference_headers, start=1):
        st.write(f"**{i}.** {h}")

st.success(
    "Consolidation completed. Cell formatting for retained rows is copied from "
    "the original source files, while the first uploaded workbook supplies the overall template."
)

st.download_button(
    "⬇️ Download Formatted Consolidated Excel",
    data=out.getvalue(),
    file_name=output_filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)

st.caption(
    "V2 preserves cell styles, fills/colors, fonts, borders, alignment, number formats, "
    "hyperlinks, comments and optional row heights for retained rows."
)
