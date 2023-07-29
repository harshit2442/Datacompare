# File Comparison Tool

This Python script is designed to compare data between two Excel files, `source.xlsx` and `target.xlsx`, and perform various comparisons to identify matches, mismatches, and missing data. The script uses the `pandas` library to read and manipulate data from the Excel files.

## Setup and Dependencies

Before running the script, ensure you have the following dependencies installed:

- Python (3.x recommended)
- pandas library (`pip install pandas`)

## Functionality

The script consists of several functions to achieve its objectives:

### 1. `read_excel()`

This function reads the data from both `source.xlsx` and `target.xlsx` files. It returns two pandas DataFrames representing the source and target data.

### 2. `count_of_fields()`

This function utilizes the `read_excel()` function to read the data. It then counts occurrences of each file name and extension in both source and target datasets. It creates a new DataFrame, `file_counts`, that holds the count information and adds an "Outcome" column to determine whether a file name and extension are matched or not between the source and target datasets.

### 3. `compare_fields()`

The main function of the script that performs the comparison. It uses the `read_excel()` and `count_of_fields()` functions to read the data and perform the initial counts.

- It checks for non-PDF file extensions in both the source and target data and creates separate DataFrames for them.
- For PDF file extensions, it compares the data in the source and target DataFrames using a common "Key" column.
- It identifies identical data rows, mismatched data rows, and missing keys between the source and target DataFrames.
- It writes the comparison results to an output Excel file named `output_file.xlsx`, separating the data into different sheets based on the comparison outcomes.

## Running the Script

To run the script, follow these steps:

1. Install the required dependencies as mentioned in the "Setup and Dependencies" section.
2. Place the `source.xlsx` and `target.xlsx` files in the same directory as the script.
3. Execute the script by running `python script.py` in the terminal or command prompt.

The script will perform the data comparison and create an `output_file.xlsx` containing the results in different sheets.

## Output

The script generates an output Excel file named `output_file.xlsx` with the following sheets:

1. **NotFound**: Contains data from the source file that couldn't be found in the target file.
2. **DataMismatch**: Contains data that exists in both source and target files but with differences.
3. **IdenticalData**: Contains data that is identical between the source and target files.
4. **TargetMismatch**: Contains data from the target file that doesn't exist in the source file.
5. **sourceOtherext**: Contains data rows from the source file with non-PDF file extensions.
6. **targetOtherext**: Contains data rows from the target file with non-PDF file extensions.
7. **FileCounts**: Provides the count of occurrences for each file name and extension, along with the outcome (matched or not).

## Note

- Make sure the column "Key" is present in both `source.xlsx` and `target.xlsx` files, as it is used to match rows between the datasets.
- Ensure the file extensions are correctly defined and present in the "File_Extension" column.
- This script assumes the first sheet in both Excel files (`Sheet1`) contains the relevant data.

Feel free to modify the script or adjust the comparison criteria based on your specific use case.
