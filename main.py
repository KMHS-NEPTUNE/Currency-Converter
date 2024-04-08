import openpyxl
import typer
import requests
import warnings

warnings.filterwarnings(action='ignore')
app = typer.Typer()

excel_list = []


def exchange_rate_data_get():
    return requests.get("https://open.er-api.com/v6/latest/JPY").json()["rates"]["KRW"]


@app.command()
def read_excel(file: str):
    count = 3
    print("Reading Excel file...")
    wb = openpyxl.load_workbook(file)
    sheet = wb.get_sheet_by_name(wb.sheetnames.pop(0))
    # sheet.cell(row=1, column=1).value
    for row in sheet.iter_rows(values_only=True):
        excel_list.append(row)
    print("Reading Excel file is done.")
    print(exchange_rate_data_get())

    exchange_rate = exchange_rate_data_get()

    for i in excel_list:
        if i == excel_list[0]:
            continue
        elif i == excel_list[1]:
            continue
        name = i[0]
        KRW = round(i[1])
        JPY = round(i[2])
        JPY_to_KRW = round(i[2] * exchange_rate)

        sheet = wb.get_sheet_by_name(wb.sheetnames.pop(1))
        # row = A, column = 1
        sheet.cell(row=count, column=1).value = i[0]
        sheet.cell(row=count, column=2).value = KRW
        sheet.cell(row=count, column=3).value = JPY_to_KRW

        sheet = wb.get_sheet_by_name(wb.sheetnames.pop(2))
        sheet.cell(row=count, column=1).value = i[0]

        if i[1] - JPY_to_KRW < 0:
            sheet.cell(row=count, column=2).value = (KRW - JPY_to_KRW) * -1
        else:
            sheet.cell(row=count, column=2).value = KRW - JPY_to_KRW

        if (i[1] - JPY_to_KRW) / i[1] * 100 < 0:
            sheet.cell(row=count, column=3).value = round((i[1] - JPY_to_KRW) / i[1] * 100 * -1)
        else:
            sheet.cell(row=count, column=3).value = round((i[1] - JPY_to_KRW) / i[1] * 100)

        count += 1
    wb.save(file)


if __name__ == "__main__":
    app()
