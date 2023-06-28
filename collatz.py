print("Starting now")
import os.path
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
    if not os.path.isfile(filename):
        wb = Workbook()
        wb.active.title = "Sheet1"
        wb.save(filename)
        wb.close()
        write_to_excel(filename, [["Number", "Steps"]]) 
    else:
        wb = load_workbook(filename)
        sheet = wb.active
        for row_data in data:
            sheet.append(row_data)
            print("Number written:", row_data[0])  # Print the number
        wb.save(filename)
        wb.close()

def collatz_to_excel(filename, start, end):
    directory = os.path.dirname(filename)
    os.makedirs(directory, exist_ok=True)

    data_to_write = []

    for num in range(start, end+1):
        steps = collatz_steps(num)
        data_to_write.append([num, steps])

        if len(data_to_write) > 0:  # Write as long as there is at least one number to write
            write_to_excel(filename, data_to_write)
            data_to_write = []

    if data_to_write:
        write_to_excel(filename, data_to_write)

    print("All numbers written!")

start = 1
end = 2**10
collatz_to_excel(f"Excels/collatz_steps {start} to {end}.xlsx", start, end)
