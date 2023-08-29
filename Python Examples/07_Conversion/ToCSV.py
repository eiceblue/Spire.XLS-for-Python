from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ToCSV.xlsx"
outputFile = "ToCSV.csv"

#create a workbook
workbook = Workbook()
#load a excel document
workbook.LoadFromFile(inputFile)
#get the first sheet
sheet = workbook.Worksheets[0]
#convert to CSV file
sheet.SaveToFile(outputFile, ",", Encoding.get_UTF8())
workbook.Dispose()


