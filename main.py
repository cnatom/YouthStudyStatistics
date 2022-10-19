from openpyxl import load_workbook, Workbook
from collections import Counter
from openpyxl.worksheet.worksheet import Worksheet


def get_counter(workbook: str) -> dict:
    result: dict = {}
    raw: Workbook = load_workbook(workbook)
    for sheet_name in raw.sheetnames:
        sheet: Worksheet = raw[sheet_name]
        cols: list = sheet["B"][1:]
        cols = list(filter(None, [item.value for item in cols]))
        cur_result: dict = Counter(cols)
        result[sheet_name] = cur_result
    return result


def export(filename: str, counter: dict):
    wb: Workbook = Workbook()
    for key, value_dict in counter.items():
        sheet: Worksheet = wb.create_sheet(key)
        sheet.cell(row=1, column=1, value="班级")
        sheet.cell(row=1, column=2, value="完成数量")
        i: int = 2
        for value_k, value_v in value_dict.items():
            sheet.cell(row=i, column=1, value=value_k)
            sheet.cell(row=i, column=2, value=value_v)
            i = i + 1
    del wb["Sheet"]
    wb.save(filename)


if __name__ == '__main__':
    counter: dict = get_counter(workbook="raw.xlsx")
    export(filename="result.xlsx", counter=counter)
