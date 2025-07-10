from spire.xls import *
from spire.xls.common import *


outputFile = "ColorsAndPalette.xlsx"

#Create a workbook
workbook = Workbook()
#Adding Orchid color to the palette at 60th index
workbook.ChangePaletteColor(Color.get_Orchid(), 60)
#Get the first sheet
sheet = workbook.Worksheets[0]
cell = sheet.Range["B2"]
cell.Text = "Welcome to use Spire.XLS"
#Set the Orchid (custom) color to the font
cell.Style.Font.Color = Color.get_Orchid()
cell.Style.Font.Size = 20
cell.AutoFitColumns()
cell.AutoFitRows()
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

