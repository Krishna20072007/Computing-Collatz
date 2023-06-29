import openpyxl
import matplotlib.pyplot as plt
import numpy as np

start = 16384
end = 42000
desired_width = 10000 * 35  # Width in pixels
desired_dpi = 80  # Adjust this value to reduce the image size

# Load the workbook
workbook = openpyxl.load_workbook(
    f'Excels/collatz_steps {start} to {end}.xlsx')

# Select the active sheet
sheet = workbook.active

# Get the values from columns A and B, skipping the header row
x_values = []
y_values = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    x_values.append(row[0])
    y_values.append(row[1])

# Convert lists to numpy arrays for faster processing
x_values = np.array(x_values)
y_values = np.array(y_values)

# Calculate the desired width in inches based on pixels and DPI
dpi = desired_dpi
desired_width_inches = desired_width / dpi

# Increase the width of the figure
fig, ax = plt.subplots(figsize=(desired_width_inches, 6), dpi=dpi)

# Plot the graph
plt.plot(x_values, y_values)
plt.xlabel('X')
plt.ylabel('Y')
plt.title('X vs Y')

# Add (x, y) coordinates as annotations
for i in range(len(x_values)):
    plt.annotate(f"({x_values[i]}, {y_values[i]})", xy=(
        x_values[i], y_values[i]), xytext=(5, 5), textcoords='offset points')

# Save the graph as an SVG file with the same name as the Excel file
svg_filename = f'collatz_steps {start} to {end}.svg'
plt.savefig(svg_filename, format='svg')

# Close the workbook
workbook.close()
print("Done")
