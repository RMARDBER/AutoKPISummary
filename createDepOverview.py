import os

import pandas as pd
from openpyxl import load_workbook

EXTRACTION_MAP_1 = {
    "SW": "A3",
    "Vehicle ID": "A4",
    "KPI": "H1",
}

SW_ID_LEN = 5 # use last 5 chars as SW identifier

COLUMNS = ["Test Name", "Test week"] + list(EXTRACTION_MAP_1.keys()) + ["Summary"]

DEP_TEST_SHEET_NAME = "SPA3_"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SOURCE_FILE = os.path.join(BASE_DIR, "data", "PKV Dependability SPA3 Test workbook.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "data", "PKV Dependability Overview.xlsx")


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
                value = float(value)
            # print(f"{key} ({cell}): {value}")
            df_out.at[sheet_name, key] = value

        print(df_out.at[sheet_name, "SW"][-SW_ID_LEN:],":",sw_to_test_week.get(df_out.at[sheet_name, "SW"][-SW_ID_LEN:]))
        df_out.at[sheet_name, "Test week"] = sw_to_test_week.get(df_out.at[sheet_name, "SW"][-SW_ID_LEN:])


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