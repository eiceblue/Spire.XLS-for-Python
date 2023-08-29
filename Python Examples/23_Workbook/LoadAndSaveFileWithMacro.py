from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/MacroSample.xls"
outputFile = "LoadAndSaveFileWithMacro.xls"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
sheet.Range["A5"].Text = "This is a simple test!"
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version97to2003)
workbook.Dispose()

