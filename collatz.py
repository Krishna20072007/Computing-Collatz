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
        write_to_excel(filename, ["Number", "Steps"])  # Pass header as a flat list
    else:
        wb = load_workbook(filename)
        sheet = wb.active
        sheet.append(data)
        print("Number written:", data[0])  # Print the number
        wb.save(filename)
        wb.close()

def collatz_to_excel(filename, start, end):
    directory = os.path.dirname(filename)
    os.makedirs(directory, exist_ok=True)

    for num in range(start, end+1):
        steps = collatz_steps(num)
        write_to_excel(filename, [num, steps])

    print("All numbers written!")

start = 1
end = 2**5
collatz_to_excel(f"Excels/collatz_steps {start} to {end}.xlsx", start, end)
