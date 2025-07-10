from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WorksheetSample2.xlsx"
outputFile = "ShowOrHideGridLine.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first and second worksheet
sheet1 = workbook.Worksheets[0]
sheet2 = workbook.Worksheets[1]
#Hide grid line in the first worksheet
sheet1.GridLinesVisible = False
#Show grid line in the first worksheet
sheet2.GridLinesVisible = True
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

