import pandas as pd
import xlsxwriter

# Define the path to your existing .xlsx file
xlsx_file = "U:/AJ Dev/TEST/Funding Transactions_1729262777420.xlsx"

# Read the Excel file into a DataFrame
df = pd.read_excel(xlsx_file)

# Define the output file where the table will be saved
output_file = "U:/AJ Dev/TEST/Funding Transactions_1729262777420.xlsx"

# Create a Pandas Excel writer using XlsxWriter as the engine
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    # Write the DataFrame to the Excel writer
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    # Access the workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # Get the number of rows and columns in the DataFrame
    max_row, max_col = df.shape

    # XlsxWriter uses 0-based indexing for rows and columns, so we add 1 for Excel's 1-based system.
    # Define the range for the table using Excel-style references (A1, B1, etc.)
    # Convert the column number into Excel column letters
    from xlsxwriter.utility import xl_col_to_name

    # The table range in Excel is from A1 to the bottom-right corner of the data
    table_range = f'A1:{xl_col_to_name(max_col - 1)}{max_row + 1}'

    # Add a table to the worksheet
    worksheet.add_table(table_range, {
        'columns': [{'header': col} for col in df.columns]
    })

# The file will be saved automatically when exiting the 'with' block
print(f"Table created and saved in {output_file}")
