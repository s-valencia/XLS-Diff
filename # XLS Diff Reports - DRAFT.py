# XLS Diff Reports - DRAFT
# Description: This script is designed to take two Excel files as input, 
#               create a duplicate of the first file with updated headers, 
#               highlight differences between the two files, 
#               and send an email with the differences file as an attachment.

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def main():
    ## Step 1: Take in two different Excel files
    excel_before_path = 'PGE_Report_Field_Maps before - Copy.xlsx'
    excel_after_path = 'PGE_Report_Field_Maps after - Copy.xlsx'

    excel_before = pd.read_excel(excel_before_path)
    excel_after = pd.read_excel(excel_after_path)


    ## Step 2: Create a duplicate of “excel_before” and modify headers
    excel_differences_report = excel_before.copy()
    excluded_headers = ["Service Item ID", "Service Item ArcGIS Online URL", "Item Type"]
    og_headers = [col for col in excel_differences_report.columns if col not in excluded_headers]
    after_headers = [f"{col} (after)" for col in og_headers]


    ## Step 3: Add a blank column after every header
    new_columns = []
    for i, column in enumerate(excel_differences_report.columns):
        new_columns.append(column)  # Add the original column
        if i == 0 or (i > 0 and column not in excluded_headers):
            new_columns.append(f"{column} (after)")  # Add a blank column header with a unique name
    
    # Create the resulting DataFrame with updated column structure
    new_df = pd.DataFrame(columns=new_columns)
    for i, column in enumerate(excel_differences_report.columns):
        new_df[column] = excel_differences_report[column]  # Copy original column data
        if i == 0 or (i > 0 and column not in excluded_headers):
            new_df[f"{column} (after)"] = ""  # Add a blank column with a unique name
    
    excel_differences_report = new_df


    ## Step 4: Create a duplicate of “excel_after” and modify headers
    excel_after2 = excel_after.copy()
    after_headers = {
        col: f"{col} (after)" if col not in excluded_headers else col for col in excel_after2.columns
    }
    excel_after2.rename(columns=after_headers, inplace=True)


    ## Step 5: Merge all data that have matching headers from “excel_after2” to “XLS Differences Report”
    for column in excel_after2.columns:
        if column in excel_differences_report.columns:
            excel_differences_report[column] = excel_after2[column]


    ## Step 6: Move excluded headers to the first three columns
    ordered_columns = excluded_headers + [col for col in excel_differences_report.columns if col not in excluded_headers]
    excel_differences_report = excel_differences_report[ordered_columns]
    
    
    ## Step 7: Analyze for differences and highlight in Excel
    output_file = "XLS Differences Report (test).xlsx"
    excel_differences_report.to_excel(output_file, index=False, engine='openpyxl')
    wb = load_workbook(output_file)
    sheet = wb.active

    # Define a fill style for highlighting
    highlight_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Start checking from columns 4 and 5, then every two columns after that
    col_index = 4

    while True:
        col1 = sheet.cell(row=1, column=col_index).column_letter
        col2 = sheet.cell(row=1, column=col_index + 1).column_letter

        # Break the loop if either header is blank
        if sheet[f"{col1}1"].value is None or sheet[f"{col2}1"].value is None:
            break

        for row in range(2, sheet.max_row + 1):  # Assuming the first row is headers
            value1 = sheet[f"{col1}{row}"].value
            value2 = sheet[f"{col2}{row}"].value

            if value1 != value2:  # Highlight mismatched cells
                sheet[f"{col1}{row}"].fill = highlight_fill
                sheet[f"{col2}{row}"].fill = highlight_fill

        col_index += 2

    # Save the updated workbook
    wb.save(output_file)

    ## Step 8: Print a brief summary and bullet points of differences found
    differences = []
    for col in excel_differences_report.columns:
        if col.endswith(" (after)"):
            before_col = col.replace(" (after)", "")
            diff_count = (excel_differences_report[col] != excel_differences_report[before_col]).sum()
            if diff_count > 0:
                differences.append(f"'{before_col}': {diff_count} differences found.")

    # Summarize the differences in a professional manner
    if differences:
        print("Summary of Differences: ")
        print("\n")
        print("\n".join([f"- {diff}" for diff in differences]))
    else:
        print("No differences were found between the two Excel files.")

if __name__ == "__main__":
    main()