import pandas as pd
import xlsxwriter

xlsx_file = "U:/AJ Dev/TEST/Funding Transactions_1729262777420.xlsx"


df = pd.read_excel(xlsx_file)


output_file = "U:/AJ Dev/TEST/Funding Transactions_1729262777420.xlsx"


with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
  
    df.to_excel(writer, sheet_name='Sheet1', index=False)

   
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

   
    max_row, max_col = df.shape

    from xlsxwriter.utility import xl_col_to_name

    
    table_range = f'A1:{xl_col_to_name(max_col - 1)}{max_row + 1}'

    
    worksheet.add_table(table_range, {
        'columns': [{'header': col} for col in df.columns]
    })

print(f"Table created and saved in {output_file}")
