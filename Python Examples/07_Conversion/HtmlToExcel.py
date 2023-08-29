from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/HtmlToExcel.html"
outputFile = "HtmlToExcel.xlsx"

#Create a workbook
workbook = Workbook()
#Load html
workbook.LoadFromHtml(inputFile)
#Save the document and launch it
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)

