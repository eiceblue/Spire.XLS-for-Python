from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Sample.xlsx"
outputFile = "ConvertTextToNubmer.xlsx"

#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
worksheet = workbook.Worksheets[0]
#Convert text string format to number format
worksheet.Range["D2:D8"].ConvertToNumber()
#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


