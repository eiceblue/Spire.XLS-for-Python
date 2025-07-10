from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_1.xlsx"
outputFile = "DuplicateCellRange.xlsx"
        
#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Copy data from source range to destination range and maintain the format.
sheet.Copy(sheet.Range["A6:F6"], sheet.Range["A16:F16"], True)
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

