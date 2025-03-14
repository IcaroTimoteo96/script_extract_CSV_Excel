import csv
from datetime import datetime
import xlsxwriter

_csv = r"./transacoes_ficticias.csv"

def extract_csv(_csv):
    columns = []
    rows = []
    copy_rows = []
    date_format = '%Y-%m-%d %H:%M:%S'

    month = int(input("Digite o número do mês (1-12) para verificar o total de transações:\n"))
    try:
        print("Lendo arquivo csv...")
        with open(_csv, encoding='utf-8',mode="r") as csvfile:
            csvreader = csv.reader(csvfile)

            columns = next(csvreader)

            for row in csvreader:
                rows.append(row)

            for row in rows:
                date_obj = datetime.strptime(row[1], date_format)
                if date_obj.month == month:
                    copy_rows.append(row)

            create_excel(columns, copy_rows, month)

    except Exception as ex:
        print(ex)

def create_excel(columns, rows, month):
    try:
        excel_col = 0

        workbook = xlsxwriter.Workbook(f"transacoes_mes_{month}.xlsx")
        worksheet = workbook.add_worksheet()
        bold = workbook.add_format({'bold': True})

        for col in columns:
            worksheet.write(0, excel_col, col, bold)
            excel_col+=1

        next_column = 0
        next_rows = 1

        for row in rows:
            for element in row:
                worksheet.write(next_rows, next_column, element)
                next_column += 1

            next_column = 0
            next_rows += 1

        workbook.close()

    except Exception as ex:
        print(ex)


extract_csv(_csv)

