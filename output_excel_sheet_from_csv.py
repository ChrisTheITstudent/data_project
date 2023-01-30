import csv
import openpyxl
from datetime import datetime, timedelta


class Record:
    def __init__(self, product, production_date, barcode, age):
        self.product = product
        self.production_date = production_date
        self.barcode = barcode
        self.age = age


filename = "c:/Users/c_mil/Desktop/Dev/data.csv"

records = []

with open(filename, "r") as file:
    reader = csv.DictReader(file)
    headers = [header.strip() for header in reader.fieldnames]
    for row in reader:
        if "product" in headers and "production date" in headers and "carton barcode" in headers:
            production_date = datetime.strptime(
                row["production date"], "%d/%m/%Y").date()
            age = (datetime.now().date() - production_date).days
            if age > 14:
                record = Record(row["product"], production_date,
                                row["carton barcode"], age)
                records.append(record)

if records:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = str(datetime.now().date())

    headers = ["Product", "Production Date", "Carton Barcode", "Product Age"]
    sheet.append(headers)

    for record in records:
        sheet.append([record.product, record.production_date,
                     record.barcode, record.age])

    table = openpyxl.worksheet.table.Table(
        ref=f"A1:{chr(ord('A') + len(headers) - 1)}{len(records) + 1}", displayName=sheet.title)
    sheet.add_table(table)

    workbook.save(f"{sheet.title}.xlsx")
else:
    print("No records found.")
