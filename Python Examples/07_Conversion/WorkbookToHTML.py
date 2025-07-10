from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WorkbookToHTML.xlsx"
outputFile = "WorkbookToHTML.html"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Convert to html
workbook.SaveToHtml(outputFile)
workbook.Dispose()

