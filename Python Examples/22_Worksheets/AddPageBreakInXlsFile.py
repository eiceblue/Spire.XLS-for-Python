from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "AddPageBreakInXlsFile.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Add page break in Excel file.
sheet.HPageBreaks.Add(sheet.Range["E4"])
sheet.VPageBreaks.Add(sheet.Range["C4"])
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

