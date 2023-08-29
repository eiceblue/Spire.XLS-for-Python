from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/WorksheetSample4.xlsx"
outputFile = "ShowTab.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Show worksheet tab
workbook.ShowTabs = True
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

