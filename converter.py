import tabula
import pandas as pd

# Path to the PDDF file
pdf_path = 'pdf/Riske 2022 Crypto Sales.pdf'

# Pages containing the table
pages = '3-26'

# Extract tables from the PDF
tables = tabula.read_pdf(pdf_path, pages=pages, multiple_tables=True)

# Merge tables into a single DataFrame
table_df = pd.concat(tables,ignore_index=True)

# Path to the Excel file
excel_output_path = 'excel/output.xlsx'

# Export DataFrame to Excel
table_df.to_excel(excel_output_path, index=False)