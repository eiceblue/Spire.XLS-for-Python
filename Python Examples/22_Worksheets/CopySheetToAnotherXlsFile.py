from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/CopySheetToAnotherXlsFile.xlsx"
outputFile = "sourceFile.xlsx"

#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Put some data into header rows (A1:A4)
for i in range(1, 6):
    sheet.Range["A" + str(i)].Text = "Header Row {0}".format(i)
    #sheet.Cells[i].Value = string.Format("Header Row {0}",i)
#Put some detail data (A5:A99)
for i in range(5, 100):
    sheet.Range["A" + str(i)].Text = "Detail Row {0}".format(i)
    #sheet.Cells[i].Value = string.Format("Detail Row {0}",i)
#Define a pagesetup object based on the first worksheet.
pageSetup = sheet.PageSetup
#The first five rows are repeated in each page. It can be seen in print preview.
pageSetup.PrintTitleRows = "$1:$5"
#Create another Workbook.
workbook1 = Workbook()
#Get the first worksheet in the book.
sheet1 = workbook1.Worksheets[0]
#Copy worksheet to destination worsheet in another Excel file.
sheet1.CopyFrom(sheet)
#Save the document
workbook1.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

