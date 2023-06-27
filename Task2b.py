import os
import xlrd
import matplotlib.pyplot as plt

# Folder path containing the Excel files, you can change this based on the path you store your excel files.
folder_path = 'E:\casimirCourseTask'

# List of Excel files to process
excel_files = ['sample1.xls', 'sample2.xls', 'sample3.xls', 'sample4.xls', 'sample5.xls', 'sample6.xls']

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
    
    # Remove file extension from the label
    label = os.path.splitext(file)[0]
    
    # Plot the data with a unique color for each file
    plt.plot(modified_Displacement, Load, label=label)

# Set labels and title
plt.xlabel('Displacement(mm)')
plt.ylabel('Load(kN)')
plt.title('Load-Displacement')

# Set x-axis limit to start from 0
plt.xlim(0)
plt.ylim(0)

# Show the legend
plt.legend()

# Save the figure as a high-quality PNG format
plt.savefig('results.png', dpi=300)

# Show the plot
plt.show()