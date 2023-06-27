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

def write_to_excel(filename, num, steps):
    wb = load_workbook(filename)
    sheet = wb.active
    row = sheet.max_row + 1
    sheet.cell(row=row, column=1).value = num
    sheet.cell(row=row, column=2).value = steps
    wb.save(filename)
    wb.close()

def collatz_to_excel(filename, start_num, max_rows, start, end, batch_size=100):
    directory = os.path.dirname(filename)
    if not os.path.exists(directory):
        os.makedirs(directory)

    if not os.path.isfile(filename):
        wb = Workbook()
        wb.active.title = "Sheet1"
        wb.save(filename)
        wb.close()

    current_row = start_num

    while current_row <= end and current_row < start_num + max_rows:
        steps = collatz_steps(current_row, start, end)
        if steps > 0:
            write_to_excel(filename, current_row, steps)
            if current_row % batch_size == 0:
                print(f"Numbers written: {current_row}")
        current_row += 1

    print("All numbers written!")

# Example usage
start = 1
end = 2**10
collatz_to_excel("Excels/collatz_steps.xlsx", start, end, start, end)
