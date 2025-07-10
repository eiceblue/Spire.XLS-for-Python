from spire.xls import *
from spire.xls.common import *


outputFile = "./Demos/UsePredefinedStyles.xlsx"

#create a workbook
workbook = Workbook()

#get the first worksheet
sheet=workbook.Worksheets[0]

#create a new style
style = workbook.Styles.Add("newStyle")
style.Font.FontName = "Calibri"
style.Font.IsBold = True
style.Font.Size = 15
style.Font.Color = Color.get_CornflowerBlue()

#get "B5" cell
range = sheet.Range["B5"]
range.Text = "Welcome to use Spire.XLS"
range.CellStyleName = style.Name
range.AutoFitColumns()

#save the file
workbook.SaveToFile(outputFile, FileFormat.Version2013)
workbook.Dispose()


