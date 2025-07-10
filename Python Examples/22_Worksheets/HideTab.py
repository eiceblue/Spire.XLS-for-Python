from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WorksheetSample2.xlsx"
outputFile = "HideTab_1.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Hide worksheet tab
workbook.ShowTabs = False
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()