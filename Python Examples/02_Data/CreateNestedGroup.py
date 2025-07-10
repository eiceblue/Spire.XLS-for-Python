from spire.xls.common import *
from spire.xls import *


outputFile = "CreateNestedGroup.xlsx"

#Create a workbook.
workbook = Workbook()

#Get the first worksheet.
sheet = workbook.Worksheets[0]

#Set the style.
style = workbook.Styles.Add("style")
style.Font.Color = Color.get_CadetBlue()
style.Font.IsBold = True

#Set the summary rows appear above detail rows.
sheet.PageSetup.IsSummaryRowBelow = False

#Insert sample data to cells.
sheet.Range["A1"].Value = "Project plan for project X"
sheet.Range["A1"].CellStyleName = style.Name

sheet.Range["A3"].Value = "Set up"
sheet.Range["A3"].CellStyleName = style.Name
sheet.Range["A4"].Value = "Task 1"
sheet.Range["A5"].Value = "Task 2"
sheet.Range["A4:A5"].BorderAround(LineStyleType.Thin)
sheet.Range["A4:A5"].BorderInside(LineStyleType.Thin)

sheet.Range["A7"].Value = "Launch"
sheet.Range["A7"].CellStyleName = style.Name
sheet.Range["A8"].Value = "Task 1"
sheet.Range["A9"].Value = "Task 2"
sheet.Range["A8:A9"].BorderAround(LineStyleType.Thin)
sheet.Range["A8:A9"].BorderInside(LineStyleType.Thin)

#Group the rows that you want to group.
sheet.GroupByRows(2, 9, False)
sheet.GroupByRows(4, 5, False)
sheet.GroupByRows(8, 9, False)

#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
