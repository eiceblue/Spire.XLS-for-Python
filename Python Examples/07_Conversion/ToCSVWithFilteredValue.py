from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/AutofilterSample.xlsx"
outputFile = "ToCSVWithFilteredValue.csv"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Convert to CSV file with filtered value
workbook.Worksheets[0].SaveToFile(outputFile, ";", False)
workbook.Dispose()

