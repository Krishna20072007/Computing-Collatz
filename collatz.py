import os
import openpyxl
import sys

add = 166440
start = 2**70+(add*10) 
end = start+add 

def collatz(x, collatz_dict):
    sequence = [x]
    while x != 1:
        if x % 2 == 0:
            x //= 2
        else:
            x = 3 * x + 1
        sequence.append(x)
        if x in collatz_dict:
            sequence.extend(collatz_dict[x][1:])
            break
    collatz_dict[sequence[0]] = sequence
    return sequence


def save_to_excel(collatz_dict, start, end):
    filename = f"Excels/collatz {start} - {end}.xlsx"

    if not os.path.exists("Excels"):
        os.makedirs("Excels")

    if not os.path.exists(filename):
        workbook = openpyxl.Workbook()
        workbook.save(filename)

    workbook = openpyxl.load_workbook(filename)
    worksheet = workbook.active

    row = worksheet.max_row + 1

    for num, sequence in collatz_dict.items():
        worksheet.cell(row=row, column=1).value = num
        for col, value in enumerate(sequence, start=2):
            worksheet.cell(row=row, column=col).value = value
        row += 1

        print(f"::warning file=collatz.py,line=45::Number {num} saved to the Excel file.")
        sys.stdout.flush()

    workbook.save(filename)


def save_step_counter(start, end):
    directory = "Steps"
    if not os.path.exists(directory):
        os.makedirs(directory)

    filename = f"{directory}/collatz {start} - {end}.txt"
    if not os.path.exists(filename):
        with open(filename, 'w') as file:
            file.write(f"Step counter for numbers {start} to {end}:\n\n")
    else:
        print("File already exists. Skipping creation.")

    for i in range(start, end + 1):
        sequence = collatz(i, {})
        steps = len(sequence) - 1
        with open(filename, 'a') as file:
            file.write(f"Number: {i}\n")
            file.write(f"Steps: {steps}\n")

        print(f"{i}")
        sys.stdout.flush()


def main():
    save_step_counter(start, end)
    collatz_dict = {}

    for i in range(start, end + 1):
        sequence = collatz(i, collatz_dict)
        del collatz_dict[i]

    save_to_excel(collatz_dict, start, end)


main()
