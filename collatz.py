import os
from openpyxl import Workbook, load_workbook

def collatz_steps(n):
    steps = 0
    while n != 1:
        if n % 2 == 0:
            n = n // 2
        else:
            n = 3 * n + 1
        steps += 1
    return steps

def write_to_excel(filename, data):
    wb = load_workbook(filename)
    sheet = wb.active
    for row_data in data:
        sheet.append(row_data)
    wb.save(filename)
    wb.close()

def collatz_to_excel(filename, start, end):
    directory = os.path.dirname(filename)
    os.makedirs(directory, exist_ok=True)

    if not os.path.isfile(filename):
        wb = Workbook()
        wb.active.title = "Sheet1"
        wb.save(filename)
        wb.close()
        write_to_excel(filename, [["Number", "Steps"]])  # Add headers to the sheet

    data_to_write = []

    for num in range(start, end+1):
        steps = collatz_steps(num)
        if steps > 0:
            data_to_write.append([num, steps])

    if data_to_write:
        write_to_excel(filename, data_to_write)

    print("All numbers written!")

# Example usage
start = 1
end = 2**20
collatz_to_excel("Excels/collatz_steps.xlsx", start-1, end)
