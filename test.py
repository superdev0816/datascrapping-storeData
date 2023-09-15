import xlsxwriter

# Create a sample XLSX file
workbook = xlsxwriter.Workbook('example.xlsx')
worksheet = workbook.add_worksheet()

# Define the hyperlink format
hyperlink_format = workbook.add_format({'color': 'blue', 'underline': 1})

# Write the hyperlink to the cell
worksheet.write_url('A1', 'https://www.google.com', hyperlink_format, 'Click here')

# Close the workbook
workbook.close()