from spire.xls import *

inputFile = "./Demos/Data/DeleteDuplicateRows.xlsx"
outputFile = "DeleteDuplicateRows_out.xlsx"

#Create a workbook.
workbook = Workbook()

#Load the file from disk.
workbook.LoadFromFile(inputFile)

#Get the first worksheet.
sheet = workbook.Worksheets.get_Item(0)

# Remove duplicate rows from the sheet
sheet.RemoveDuplicates()

#Save to file.
workbook.SaveToFile(outputFile, FileFormat.Version2013)
workbook.Dispose()