import os

import pandas as pd
from openpyxl import load_workbook

EXTRACTION_MAP_1 = {
    "SW": "A3",
    "Vehicle ID": "A4",
    "KPI": "H1",
}

SW_ID_LEN = 5 # use last 5 chars as SW identifier

COLUMNS = ["Test Name", "Test week"] + \
    list(EXTRACTION_MAP_1.keys()) + \
    ["Severity 5: Critical",
    "Severity 4: Major",
    "Severity 3: Moderate","Summary"]

DEP_TEST_SHEET_NAME = "SPA3_"
LOOP1_IDX = 10 # column K (index 10) is where probability values start in the sheet

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SOURCE_FILE = os.path.join(BASE_DIR, "data", "PKV Dependability SPA3 Test workbook.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "data", "PKV Dependability Overview.xlsx")

SEVERITY_LABELS = [
    "Severity 5: Critical",
    "Severity 4: Major",
    "Severity 3: Moderate",
    "Severity 2: Minor",
    "Severity 1: Low",
    "Analysis",
    "New",
]

def to_numeric_score(value):
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        normalized = value.strip()
        if not normalized:
            return None
        if normalized in {"OK", "NT"}:
            return 0.0
        # normalized = normalized.replace(",", ".")
        try:
            return float(normalized)
        except ValueError:
            return None
    return None


def calc_row_probability(row):
    # Count probability entries
    values = []
    for cell in row[LOOP1_IDX:]: #start from column K (index 10)
        numeric_value = to_numeric_score(cell)
        if numeric_value is not None:
            values.append(numeric_value)

    # prob = sum(values) / len(values) if values else None
    return values


def count_arts_per_severity(ws):
    """Count ART entries under each severity/analysis/new section in column A."""
    counts = {}
    probability = {}
    current_label = None

    for row in ws.iter_rows(min_col=1, max_col=ws.max_column, values_only=True):
        cell_value = row[0]
        if cell_value in SEVERITY_LABELS:
            current_label = cell_value
            counts[current_label] = 0
            probability[current_label] = []

        # elif current_label is not None and isinstance(cell_value, str) and cell_value.startswith("ART"):
        elif current_label is not None and cell_value is not None and cell_value != '.':
            counts[current_label] += 1
            probability[current_label] += calc_row_probability(row)
        
    return counts, probability

def remove_prefix(value, prefix):
    if isinstance(value, str) and value.startswith(prefix):
        return value[len(prefix):]
    return value

wb = load_workbook(SOURCE_FILE, data_only=True)

print("Sheet names:", wb.sheetnames)

# Load Test week lookup from Länkar sheet
df_lankar = pd.read_excel(SOURCE_FILE, sheet_name="Länkar", header=0)
df_lankar["SW label"] = df_lankar["SW label"].astype("string").str[-SW_ID_LEN:]
sw_to_test_week = dict(zip(df_lankar["SW label"].to_list(), df_lankar["Test week"]))

print(df_lankar["SW label"])

df_out = pd.DataFrame(columns=COLUMNS)

for sheet_name in wb.sheetnames:
    if sheet_name[0:len(DEP_TEST_SHEET_NAME)] == DEP_TEST_SHEET_NAME:
        ws = wb[sheet_name]
        df_out.at[sheet_name, "Test Name"] = sheet_name
        print(f"Processing sheet: {sheet_name}")
        for key, cell in EXTRACTION_MAP_1.items():
            value = ws[cell].value
            if key == "SW":
                value = remove_prefix(value, "SW under test: ")
            if key == "Vehicle ID":
                value = remove_prefix(value, "Vehicle ID: ")
            if key == "KPI" and value is not None:
                value = str(value)
            # print(f"{key} ({cell}): {value}")
            df_out.at[sheet_name, key] = value

        print(df_out.at[sheet_name, "SW"][-SW_ID_LEN:],":",sw_to_test_week.get(df_out.at[sheet_name, "SW"][-SW_ID_LEN:]))
        df_out.at[sheet_name, "Test week"] = sw_to_test_week.get(df_out.at[sheet_name, "SW"][-SW_ID_LEN:])

        # Count issues per severity level
        counts, probability = count_arts_per_severity(ws)
        for label in SEVERITY_LABELS:
            if label in df_out.columns:
                prob_list = probability.get(label)
                # print(label, "prob list:", prob_list)
                avg_prob = sum(prob_list) / len(prob_list) if prob_list else 0
                # print(prob_list)
                # print(avg_prob)
                df_out.at[sheet_name, label] = "Count: " + str(counts.get(label, 0)) + "\t\t (avg prob: " + str(round(avg_prob,1)) + ")"

# for row in ws.iter_rows(min_col=1, max_col=1, values_only=True):
#     print(row)

# print(ws[15].value)

def create_excel(df):
    df.to_excel(OUTPUT_FILE, sheet_name="Overview", index=False)

    wb = load_workbook(OUTPUT_FILE)
    ws = wb.active

    # Freeze header row
    ws.freeze_panes = "A2"

    # Auto-fit columns
    for column in ws.columns:
        ws.column_dimensions[column[0].column_letter].width = 30

    ws.auto_filter.ref = ws.dimensions

    wb.save(OUTPUT_FILE)
    wb.close()

def populate_excel(df):
    pass


create_excel(df_out)
os.startfile(OUTPUT_FILE)



# def excel_to_py_index(cell_reference):
#     row_index, column_index = coordinate_to_tuple(cell_reference)
#     return row_index-2, column_index-1

# print("A3:", excel_to_py_index("A3"))  # Should print (2, 0)

# sheets = pd.read_excel(SOURCE_FILE, sheet_name=None)

# print(sheets.keys())

# for sheet_name, df in sheets.items():
#     if sheet_name[0:len(DEP_TEST_SHEET_NAME)] == DEP_TEST_SHEET_NAME:
#         print(f"Processing sheet: {sheet_name}")
#         print(df.iat[excel_to_py_index("F8")])
#         for key, cell in EXTRACTION_MAP_1.items():
#             row_index, column_index = excel_to_py_index(cell)
#             value = df.iat[row_index, column_index]
#             print(f"{key} ({cell}): {value}")