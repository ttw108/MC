import xlwings as xw

# Create a new workbook
wb = xw.Book()

# Select the first sheet
sht = wb.sheets[0]

# Write some data in column A
sht.range('A1').value = ['A', 'B', 'C', 'D', 'E', 'F']

# Apply color scale formatting to the range
sht.range('A2:A7').color_scale(min_color='FF0000', max_color='00FF00')

# Save and close the workbook
wb.save('test.xlsx')
wb.close()