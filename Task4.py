import os
import xlrd
import numpy as np
import matplotlib.pyplot as plt

# Folder path containing the Excel files, you can change this based on the path you store your excel files.
folder_path = 'E:\casimirCourseTask'

# List of Excel files to process
excel_files = ['sample1.xls', 'sample2.xls', 'sample3.xls', 'sample4.xls', 'sample5.xls', 'sample6.xls']

# Initialize empty lists for storing all new_x and new_y arrays
all_new_x = []
all_new_y = []

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
    
    # Perform linear interpolation to get modified Displacement and corresponding Load
    x = modified_Displacement
    y = Load
    new_x = np.arange(0, 2.51, 0.01)  # Create array from 0 to 2.5 with 0.01 step size
    new_y = np.interp(new_x, x, y)  # Perform linear interpolation
    
    # Append new_x and new_y to the list of all arrays
    all_new_x.append(new_x)
    all_new_y.append(new_y)
    
    # Plot the interpolated data with black dashed line (thin)
    plt.plot(new_x, new_y, label=label, linestyle='--', linewidth=1, color='black')

# Compute the average of all_new_x and all_new_y
avg_x = np.mean(all_new_x, axis=0)
avg_y = np.mean(all_new_y, axis=0)

# Plot the average curve with red solid line (thick)
plt.plot(avg_x, avg_y, label='Average', linewidth=2, color='red')

# Fill the area between the min and max curves with a shaded color
min_y = np.min(all_new_y, axis=0)
max_y = np.max(all_new_y, axis=0)
plt.fill_between(new_x, min_y, max_y, color='gray', alpha=0.2)

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
plt.savefig('AveragingLoadDisplacement.png', dpi=300)

# Show the plot
plt.show()
