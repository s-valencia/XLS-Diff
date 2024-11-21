# XLS Diff Reports - DRAFT
# Description: This script is designed to take two Excel files as input, 
#               create a duplicate of the first file with updated headers, highlight differences between the two files, 
#               and send an email with the differences file as an attachment.


import pandas as pd
import os
from openpyxl import Workbook, load_workbook

def highlight_differences(df1, df2):
    """
    Function to highlight differences between two dataframes.
    Returns a DataFrame highlighting where differences occur.
    """
    diff = df1.compare(df2, align_axis=1, keep_shape=True, keep_equal=True)
    return diff

def main():
    # Step 1: Take two Excel files as input
    excel_1_path = input("Enter the path to the first Excel file: ")
    excel_2_path = input("Enter the path to the second Excel file: ")
    excel_1 = pd.read_excel(excel_1_path)
    excel_2 = pd.read_excel(excel_2_path)

    # Step 2: Create a duplicate of excel_1 with updated headers
    excluded_headers = ["Service Item ID", "Service Item ArcGIS Online URL", "Item Type"]
    new_headers = {col: f"{col}_(before)" if col not in excluded_headers else col for col in excel_1.columns}
    excel_1_duplicate = excel_1.rename(columns=new_headers)
    excel_1_duplicate_path = "excel_1_duplicate.xlsx"
    excel_1_duplicate.to_excel(excel_1_duplicate_path, index=False)
    print(f"Duplicate Excel file created at: {excel_1_duplicate_path}")

    # Step 3: Create a list of headers in excel_1_duplicate excluding specific columns
    excel_1_headers = [col for col in excel_1_duplicate.columns if col not in excluded_headers]

    # Step 4: Highlight differences between excel_1 and excel_2
    differences = highlight_differences(excel_1, excel_2)
    differences_path = "Differences_PGE_Reports_Field_Maps.xlsx"
    differences.to_excel(differences_path, index=False)
    print(f"Differences file created at: {differences_path}")

    # Step 5: Create a list of headers from excel_1 excluding specific columns
    header_fields = [col for col in excel_1.columns if col not in excluded_headers]

    # Step 6: Add duplicate headers with “_(before)” appended to excel_1
    for header in header_fields:
        excel_1[f"{header}_(before)"] = excel_1[header]

    excel_1_updated_path = "excel_1_with_duplicate_headers.xlsx"
    excel_1.to_excel(excel_1_updated_path, index=False)
    print(f"Updated Excel file with duplicate headers created at: {excel_1_updated_path}")

if __name__ == "__main__":
    main()