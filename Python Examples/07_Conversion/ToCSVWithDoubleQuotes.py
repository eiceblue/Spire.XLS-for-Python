from spire.xls.common import *
from spire.xls import *

inputFile = "Data/ToCSV.xlsx"
outputFile = "ToCSVWithDoubleQuotes.csv"

#Create a workbook.
workbook = Workbook()

# Load the Excel document from disk
workbook.LoadFromFile(inputFile)

# Setting the last parameter, addQuotationForStringValue, to true means that double quotes will be added.
workbook.SaveToFile(outputFile, ",",True)

workbook.Dispose()
