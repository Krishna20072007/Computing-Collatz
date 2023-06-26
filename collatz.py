import os.path
from openpyxl import Workbook, load_workbook

def collatz_steps(n, start, end):
    steps = 0
    while n != 1 and n >= start and n <= end:
        if n % 2 == 0:
            n = n // 2
        else:
            n = 3 * n + 1
        steps += 1
    return steps

def write_to_excel(filename, data):
    wb = load_workbook(filename)
    sheet = wb.active
    row = 1
    while sheet.cell(row=row, column=1).value is not None:
        row += 1
    sheet.cell(row=row, column=1).value = data[0]
    sheet.cell(row=row, column=2).value = data[1]
    wb.save(filename)
    wb.close()

def collatz_to_excel(filename, start_num, max_rows, start, end):
    directory = os.path.dirname(filename)
    if not os.path.exists(directory):
        os.makedirs(directory)

    if not os.path.isfile(filename):
        wb = Workbook()
        wb.active.title = "Sheet1"
        wb.save(filename)
        wb.close()

    wb = load_workbook(filename)
    sheet = wb.active
    current_row = sheet.max_row + 1

    for num in range(start_num, start_num + max_rows):
        steps = collatz_steps(num, start, end)
        if steps > 0:
            write_to_excel(filename, (num, steps))
            current_row += 1

    wb.save(filename)
    wb.close()

# Example usage
start = 1
end = 100000
collatz_to_excel("Excels/collatz_steps.xlsx", 1, 100, start, end)
