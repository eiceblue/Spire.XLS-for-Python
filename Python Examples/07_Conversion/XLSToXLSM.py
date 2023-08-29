from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/MacroSample.xls"
outputFile = "XLSToXLSM.xlsm"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Convert to xlsm
workbook.SaveToFile(outputFile, ExcelVersion.Version2007)
workbook.Dispose()

