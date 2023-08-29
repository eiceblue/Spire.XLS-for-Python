from spire.xls import *
from spire.common import *


outputFile = "DifferentHeaderFooterOnFirstPage.xlsx"

#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
sheet.Range["A1"].Text = "Hello World"
sheet.Range["F30"].Text = "Hello World"
sheet.Range["G150"].Text = "Hello World"
#Set the value to show the headers/footers for first page are different from the other pages.
sheet.PageSetup.DifferentFirst = 1
#Set the header and footer for the first page.
sheet.PageSetup.FirstHeaderString = "Different First page"
sheet.PageSetup.FirstFooterString = "Different First footer"
#Set the other pages' header and footer. 
sheet.PageSetup.LeftHeader = "Demo of Spire.XLS"
sheet.PageSetup.CenterFooter = "Footer by Spire.XLS"
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
