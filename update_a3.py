
"""
script for updates "A3" metric on generator registration
"""

import emeta
import dotenv
import os
import openpyxl

# A3_EXCEL_PATH = r"C:\Users\dgraha01\Scripts\test_a3_copy.xlsx"
A3_EXCEL_PATH = r"C:\Users\dgraha01\OneDrive - Environmental Protection Agency (EPA)\e-Manifest Teams\e-Manifest A3\FY 2022 registration tracking sheet.xlsx"

MONTHS = {
    "2021-10": "B2",
    "2021-11": "B3",
    "2021-12": "B4",
    "2022-01": "B5",
    "2022-02": "B6",
    "2022-03": "B7",
    "2022-04": "B8",
    "2022-05": "B9",
    "2022-06": "B10",
    "2022-07": "B11",
    "2022-08": "B12",
    "2022-09": "B13",
}


def main():
    metabase_auth()
    manifest_data = get_manifest_count()
    write_results(manifest_data)


def metabase_auth():
    if os.path.exists("C:\\Users\\dgraha01\\.env"):
        dotenv.load_dotenv("C:\\Users\\dgraha01\\.env")
        emeta.authenticate()
    else:
        print("HOME/.env not found")
        os._exit(1)


def get_manifest_count():
    resp = emeta.get_query("4532", "json")
    return resp_to_dict(resp)


def resp_to_dict(resp: list ):
    results ={}
    for i in resp:
        results[i["CREATED_DATE_MONTH"]] = i["COUNT"]
    return results


def write_results(data):
    wb = openpyxl.load_workbook(A3_EXCEL_PATH)
    ws = wb.active
    for i in data:
        manifest_cell = MONTHS[i]
        ws[manifest_cell] = data[i]
    wb.save(A3_EXCEL_PATH)


main()
