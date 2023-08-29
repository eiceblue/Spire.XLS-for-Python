from spire.xls import *
from spire.common import *


outputFile = "UseExplicitLineBreaks.xlsx"

#Create a workbook
workbook = Workbook()
#Get the first default worksheet
sheet1 = workbook.Worksheets[0]
#Specify a cell range
c5 = sheet1.Range["C5"]
#Set the cell width for specified range
sheet1.SetColumnWidth(c5.Column, 70)
#Put the string value with explicit line breaks
c5.Value = "Spire.XLS for .NET is a professional Excel .NET API\n that can be used to create, read, \nwrite, convert and print Excel files in any type \nof .NET(C#, VB.NET, ASP.NET, .NET Core) application. \nSpire.XLS for .NET offers object model\n Excel API for speeding up Excel programming in .NET platform -\n create new Excel documents from template, edit existing \nExcel documents and \nconvert Excel files."
#Set Text wrap
c5.IsWrapText = True
#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

