from spire.xls.common import *
from spire.xls import *


inputFile = "./Demos/Data/Template_Xls_3.xlsx"
outputFile = "ExpandAndCollapseGroups.xlsx"

#Create a workbook.
workbook = Workbook()

#Load the file from disk.
workbook.LoadFromFile(inputFile)

#Get the first worksheet.
sheet = workbook.Worksheets[0]

#Expand the grouped rows with ExpandCollapseFlags set to expand parent
sheet.Range["A16:G19"].ExpandGroup(GroupByType.ByRows, ExpandCollapseFlags.ExpandParent)

#Collapse the grouped rows
sheet.Range["A10:G12"].CollapseGroup(GroupByType.ByRows)
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

