import streamlit as st
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from copy import copy
import re
import io


# ============================================================
# PAGE
# ============================================================

st.set_page_config(
    page_title="6thsense Vardaan Auto Color",
    layout="wide"
)

st.title("6thsense Vardaan Auto Color")


# ============================================================
# HELPERS
# ============================================================

def clean_text(value):
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\u00a0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\u200c", "")
    text = text.replace("\u200d", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip().lower()


def get_number(value):
    if value is None:
        return None

    if isinstance(value, bool):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    text = text.replace(",", "").replace("%", "")

    try:
        return float(text)
    except:
        return None


def get_parentheses_number(value):
    if value is None:
        return None

    text = str(value)

    match = re.search(
        r"\(\s*([-+]?\d+(?:\.\d+)?)\s*\)",
        text
    )

    if match:
        try:
            return float(match.group(1))
        except:
            return None

    return None


# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx", "xlsm"]
)


if uploaded_file is None:
    st.info("Upload your master Excel file to continue.")
    st.stop()


# ============================================================
# LOAD ORIGINAL WORKBOOK
# ============================================================

file_bytes = uploaded_file.getvalue()

value_wb = openpyxl.load_workbook(
    io.BytesIO(file_bytes),
    data_only=True
)

format_wb = openpyxl.load_workbook(
    io.BytesIO(file_bytes),
    data_only=False
)


# ============================================================
# FIND SUMMARY SHEET
# ============================================================

if "summary" not in value_wb.sheetnames:
    st.error("Sheet named 'summary' was not found.")
    st.stop()

value_ws = value_wb["summary"]
format_ws = format_wb["summary"]


# ============================================================
# CREATE OUTPUT WORKBOOK
# ============================================================

out_wb = Workbook()

out_ws = out_wb.active
out_ws.title = "AutoGreen"


# ============================================================
# COPY VALUES + FORMATTING
# ============================================================

for row in value_ws.iter_rows():

    for value_cell in row:

        r = value_cell.row
        c = value_cell.column

        out_cell = out_ws.cell(
            row=r,
            column=c,
            value=value_cell.value
        )

        source_cell = format_ws.cell(
            row=r,
            column=c
        )

        # Font
        if source_cell.has_style:
            out_cell.font = copy(source_cell.font)
            out_cell.fill = copy(source_cell.fill)
            out_cell.border = copy(source_cell.border)
            out_cell.alignment = copy(source_cell.alignment)
            out_cell.protection = copy(source_cell.protection)

        # Number format
        out_cell.number_format = source_cell.number_format


# ============================================================
# COPY COLUMN WIDTHS
# ============================================================

for key, dimension in format_ws.column_dimensions.items():

    out_ws.column_dimensions[key].width = dimension.width
    out_ws.column_dimensions[key].hidden = dimension.hidden


# ============================================================
# COPY ROW HEIGHTS
# ============================================================

for key, dimension in format_ws.row_dimensions.items():

    out_ws.row_dimensions[key].height = dimension.height
    out_ws.row_dimensions[key].hidden = dimension.hidden


# ============================================================
# COPY MERGED CELLS
# ============================================================

for merged_range in format_ws.merged_cells.ranges:
    out_ws.merge_cells(str(merged_range))


# ============================================================
# CREATE AUTOBLUE SHEET FROM THE SAME ORIGINAL DATA/FORMATTING
# ============================================================

auto_blue_ws = out_wb.copy_worksheet(out_ws)
auto_blue_ws.title = "AutoBlue"


# ============================================================
# FIND HEADER ROW
# ============================================================

header_row = None

for row in range(1, out_ws.max_row + 1):

    value = out_ws.cell(row, 1).value

    if clean_text(value) == "symbol":
        header_row = row
        break


if header_row is None:
    st.error("Could not find header row containing 'Symbol' in Column A.")
    st.stop()


# ============================================================
# AUTOBLUE SHEET — BLUE CONDITIONS
# ============================================================

# AutoBlue is intentionally independent from AutoGreen.
# It starts from the original workbook values/formatting and applies
# only the AutoBlue blue-color conditions specified by the user.

auto_blue_header_row = None

for row in range(1, auto_blue_ws.max_row + 1):
    if clean_text(auto_blue_ws.cell(row, 1).value) == "symbol":
        auto_blue_header_row = row
        break

if auto_blue_header_row is not None:

    auto_blue_fill = PatternFill(fill_type="solid", fgColor="ADD8E6")
    auto_blue_stock_fill = PatternFill(fill_type="solid", fgColor="5B9BD5")
    auto_blue_new_stock_fill = PatternFill(fill_type="solid", fgColor="17365D")
    auto_blue_new_condition_rows = set()
    auto_blue_matched_rows = set()
    auto_blue_condition_counts = {}
    auto_blue_new_match_rows = set()


    def make_auto_blue(row, col):
        # Matching condition cell only.
        auto_blue_ws.cell(row, col).fill = copy(auto_blue_fill)

    def mark_auto_blue_stock(row, new_condition=False):
        # Column A is coloured only when this stock actually matches
        # at least one AutoBlue condition.
        if new_condition:
            auto_blue_ws.cell(row, 1).fill = copy(auto_blue_new_stock_fill)
        else:
            auto_blue_ws.cell(row, 1).fill = copy(auto_blue_stock_fill)

    def count_auto_blue_condition(row):
        auto_blue_condition_counts[row] = auto_blue_condition_counts.get(row, 0) + 1

    def normalize_auto_blue_heading(value):
        text = clean_text(value)
        return re.sub(r"[^a-z0-9]+", "", text)

    def find_auto_blue_col(target_heading):
        target = normalize_auto_blue_heading(target_heading)
        for col in range(1, auto_blue_ws.max_column + 1):
            heading = normalize_auto_blue_heading(
                auto_blue_ws.cell(auto_blue_header_row, col).value
            )
            if heading == target:
                return col
        return None

    # Sum I < -4
    col = find_auto_blue_col("sum i")
    if col:
        for row in range(auto_blue_header_row + 1, auto_blue_ws.max_row + 1):
            number = get_number(auto_blue_ws.cell(row, col).value)
            if number is not None and number < -4:
                make_auto_blue(row, col)

    # 16> C-B / Avg.4 — parentheses value < 0.50
    col = find_auto_blue_col("16> c-b / avg.4")
    if col:
        for row in range(auto_blue_header_row + 1, auto_blue_ws.max_row + 1):
            value = get_parentheses_number(auto_blue_ws.cell(row, col).value)
            if value is not None and value < 0.50:
                make_auto_blue(row, col)

    # 16< D-B / Avg.4 — parentheses value < -1 and < number after 16<
    col = find_auto_blue_col("16< d-b / avg.4")
    if col:
        for row in range(auto_blue_header_row + 1, auto_blue_ws.max_row + 1):
            cell_value = auto_blue_ws.cell(row, col).value
            if cell_value is None:
                continue
            text = str(cell_value)
            first_match = re.search(r"16<\s*([-+]?\d+(?:\.\d+)?)", text, re.IGNORECASE)
            parent_value = get_parentheses_number(cell_value)
            if first_match and parent_value is not None:
                try:
                    changing_value = float(first_match.group(1))
                except:
                    continue
                if parent_value < -1 and parent_value < changing_value:
                    make_auto_blue(row, col)

    # Sum O2H.10 — value below average
    col = find_auto_blue_col("sum o2h.10")
    if col:
        values = []
        for row in range(auto_blue_header_row + 1, auto_blue_ws.max_row + 1):
            number = get_number(auto_blue_ws.cell(row, col).value)
            if number is not None:
                values.append(number)
        if values:
            average_value = sum(values) / len(values)
            for row in range(auto_blue_header_row + 1, auto_blue_ws.max_row + 1):
                number = get_number(auto_blue_ws.cell(row, col).value)
                if number is not None and number < average_value:
                    make_auto_blue(row, col)

    # Sum O2L.10 — value below average
    col = find_auto_blue_col("sum o2l.10")
    if col:
        values = []
        for row in range(auto_blue_header_row + 1, auto_blue_ws.max_row + 1):
            number = get_number(auto_blue_ws.cell(row, col).value)
            if number is not None:
                values.append(number)
        if values:
            average_value = sum(values) / len(values)
            for row in range(auto_blue_header_row + 1, auto_blue_ws.max_row + 1):
                number = get_number(auto_blue_ws.cell(row, col).value)
                if number is not None and number < average_value:
                    make_auto_blue(row, col)

    # NEW CONDITION 1: BOTH first 2 dated O2L columns must be < -1.5
    dated_o2l_columns = []
    for col in range(1, auto_blue_ws.max_column + 1):
        heading = normalize_auto_blue_heading(
            auto_blue_ws.cell(auto_blue_header_row, col).value
        )
        if "o2l" in heading and heading != normalize_auto_blue_heading("sum o2l.10"):
            dated_o2l_columns.append(col)

    first_two_o2l = dated_o2l_columns[:2]
    if len(first_two_o2l) >= 2:
        for row in range(auto_blue_header_row + 1, auto_blue_ws.max_row + 1):
            values = [
                get_number(auto_blue_ws.cell(row, col).value)
                for col in first_two_o2l
            ]
            if all(v is not None and v < -1.5 for v in values):
                for col in first_two_o2l:
                    make_auto_blue(row, col)
                    count_auto_blue_condition(row)
                auto_blue_new_condition_rows.add(row)


    # NEW CONDITION 2: 10 "-ve" — first number before + must be > 3
    negative_col = find_auto_blue_col('10 "-ve"')
    if negative_col:
        for row in range(auto_blue_header_row + 1, auto_blue_ws.max_row + 1):
            value = auto_blue_ws.cell(row, negative_col).value
            if value is None:
                continue
            match = re.match(r"^\s*(\d+)\s*\+", str(value).strip())
            if match and int(match.group(1)) > 3:
                make_auto_blue(row, negative_col)
                count_auto_blue_condition(row)
                auto_blue_new_condition_rows.add(row)

    # NEW CONDITION 3: %Chg.1, %Chg.2 and %Chg.3 must ALL be < 0
    pct1_col = find_auto_blue_col("%chg.1")
    pct2_col = find_auto_blue_col("%chg.2")
    pct3_col = find_auto_blue_col("%chg.3")

    if pct1_col and pct2_col and pct3_col:
        for row in range(auto_blue_header_row + 1, auto_blue_ws.max_row + 1):
            pct1 = get_number(auto_blue_ws.cell(row, pct1_col).value)
            pct2 = get_number(auto_blue_ws.cell(row, pct2_col).value)
            pct3 = get_number(auto_blue_ws.cell(row, pct3_col).value)
            if pct1 is not None and pct1 < 0 and pct2 is not None and pct2 < 0 and pct3 is not None and pct3 < 0:
                make_auto_blue(row, pct1_col)
                make_auto_blue(row, pct2_col)
                make_auto_blue(row, pct3_col)
                auto_blue_condition_counts[row] = auto_blue_condition_counts.get(row, 0) + 1
                auto_blue_new_condition_rows.add(row)

    # A stock qualifies for the NEW-condition result only when ALL THREE
    # NEW conditions are satisfied in the same row.
    for row in range(auto_blue_header_row + 1, auto_blue_ws.max_row + 1):
        o2l_ok = False
        if len(first_two_o2l) >= 2:
            o2l_ok = all(
                (lambda v: v is not None and v < -1.5)(
                    get_number(auto_blue_ws.cell(row, col).value)
                ) for col in first_two_o2l
            )
        neg_ok = False
        if negative_col:
            value = auto_blue_ws.cell(row, negative_col).value
            match = re.match(r"^\\s*(\\d+)\\s*\\+", str(value).strip()) if value is not None else None
            neg_ok = bool(match and int(match.group(1)) > 3)
        pct_ok = False
        if pct1_col and pct2_col and pct3_col:
            vals = [get_number(auto_blue_ws.cell(row, col).value) for col in (pct1_col, pct2_col, pct3_col)]
            pct_ok = all(v is not None and v < 0 for v in vals)
        if o2l_ok and neg_ok and pct_ok:
            auto_blue_new_match_rows.add(row)
    new_matching_rows = list(auto_blue_new_match_rows)
    max_new_count = max(
        (auto_blue_condition_counts.get(row, 0) for row in new_matching_rows),
        default=0
    )

    for row in range(auto_blue_header_row + 1, auto_blue_ws.max_row + 1):
        if row in new_matching_rows and auto_blue_condition_counts.get(row, 0) == max_new_count:
            auto_blue_ws.cell(row, 1).fill = copy(auto_blue_new_stock_fill)
        else:
            auto_blue_ws.cell(row, 1).fill = PatternFill(fill_type=None)

