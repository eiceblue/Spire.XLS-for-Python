from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/CSVSample.csv"
outputFile = "CSVToPDF.pdf"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile, ",", 1, 1)
#Set the SheetFitToPage property as true
workbook.ConverterSetting.SheetFitToPage = True
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Autofit a column if the characters in the column exceed column width
i = 1
while i < sheet.Columns.Length:
    sheet.AutoFitColumn(i)
    i += 1
workbook.SaveToFile(outputFile, FileFormat.PDF)
workbook.Dispose()
