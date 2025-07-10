from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/AutofilterBlank.xlsx"
outputFile = "AutofilterNonBlank.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Match the non blank data
sheet.AutoFilters.MatchNonBlanks(0)
#Filter
sheet.AutoFilters.Filter()
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

