from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SampleB_2.xlsx"
outputFile = "ApplyBuiltInStyles.xlsx"

workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Apply title style
sheet.Range["A1:J1"].BuiltInStyle = BuiltInStyles.Title
#Save and launch result file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

