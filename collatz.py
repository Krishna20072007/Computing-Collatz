import openpyxl

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
    wb = openpyxl.load_workbook(filename)
    sheet = wb.active
    row = 1
    while sheet.cell(row=row, column=1).value is not None:
        row += 1
    sheet.cell(row=row, column=1).value = data[0]
    sheet.cell(row=row, column=2).value = data[1]
    wb.save(filename)

def collatz_to_excel(filename, start_num, max_rows):
    try:
        wb = openpyxl.load_workbook(filename)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        wb.active.title = "Sheet1"

    sheet = wb.active
    current_row = 1
    while sheet.cell(row=current_row, column=1).value is not None:
        current_row += 1

    for num in range(start_num, start_num + max_rows):
        steps = collatz_steps(num)
        write_to_excel(filename, (num, steps))
        current_row += 1

    wb.save(filename)
    wb.close()

# Example usage
collatz_to_excel("collatz_steps.xlsx", 1, 1048576)
