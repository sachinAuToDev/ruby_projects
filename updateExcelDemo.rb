require 'win32ole'

excel = WIN32OLE.new('Excel.Application') # Start Excel application
workbook = excel.Workbooks.Open('C:/Users/sachin/RubymineProjects/updateExcelUsingLibraryWeCreatedBefore/Book1.xlsx') # Open your Excel file
worksheet = workbook.Worksheets(1) # Select the first worksheet (index starts from 1)

# Update cell A1 with new value
worksheet.Cells(1, 1).Value = "Sachins"

workbook.Save # Save the changes
workbook.Close # Close the workbook
excel.Quit # Quit Excel application
