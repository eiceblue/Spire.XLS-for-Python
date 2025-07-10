from spire.xls.common import *
from spire.xls import *

outputFile = "InsertHtmlStringIntoCell.xlsx"

#Create a workbook
workbook = Workbook()
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Insert Html String in range A1
htmlCode = "<div>first line<strong>second line</strong>third line</div>"
range = sheet["A1"]
range.HtmlString = htmlCode
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()