import os
import xlrd
import matplotlib.pyplot as plt
import numpy as np

# Folder path containing the Excel files
folder_path = os.getcwd()

# List of Excel files to process
excel_files = ['sample1.xls', 'sample2.xls', 'sample3.xls', 'sample4.xls', 'sample5.xls', 'sample6.xls']

# Function to calculate the area under a curve using the trapezoidal rule
def calculate_area(x, y):
    area = np.trapz(y, x)
    return area

# List to store the areas
areas = []

# Process each Excel file
for file in excel_files:
    # Excel file path
    file_path = os.path.join(folder_path, file)

    # Open the Excel file
    workbook = xlrd.open_workbook(file_path)

    # Get the first worksheet
    worksheet = workbook.sheet_by_index(0)

    # Extract data from the 5th and 6th columns
    Load = []
    Displacement = []

    for row in range(1, worksheet.nrows):
        Load.append(worksheet.cell_value(row, 4))
        Displacement.append(worksheet.cell_value(row, 5))

    # Subtract the first value from the Displacement column
    first_value = Displacement[0]
    modified_Displacement = [value - first_value for value in Displacement]

    # Calculate the area under the curve
    area = calculate_area(modified_Displacement, Load)

    # Add the area to the list
    areas.append(area)

# List of labels for the x-axis
labels = ['Sample1', 'Sample2', 'Sample3', 'Sample4', 'Sample5', 'Sample6']

# Plotting the bar chart
plt.bar(labels, areas, edgecolor='black')

# Add the area values on top of the bars
for i, area in enumerate(areas):
    plt.text(i, area, f'{area:.2f}', ha='center', va='bottom')

# Set labels and title
plt.xlabel('Samples')
plt.ylabel('Area')
plt.title('Areas of Curves')

# Show the plot
plt.show()
