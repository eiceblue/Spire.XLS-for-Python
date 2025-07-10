from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/FontStyles.xlsx"
outputFile = "FontStyles.xlsx"

#Create a Workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Set font style
sheet.Range["B1"].Style.Font.FontName = "Comic Sans MS"
sheet.Range["B2:D2"].Style.Font.FontName = "Corbel"
sheet.Range["B3:D7"].Style.Font.FontName = "Aleo"
#Set font size
sheet.Range["B1"].Style.Font.Size = 45
sheet.Range["B2:D3"].Style.Font.Size = 25
sheet.Range["B3:D7"].Style.Font.Size = 12
#Set excel cell data to be bold
sheet.Range["B2:D2"].Style.Font.IsBold = True
#Set excel cell data to be underline
sheet.Range["B3:B7"].Style.Font.Underline = FontUnderlineType.Single
#set excel cell data color
sheet.Range["B1"].Style.Font.Color = Color.get_CornflowerBlue()
sheet.Range["B2:D2"].Style.Font.Color = Color.get_CadetBlue()
sheet.Range["B3:D7"].Style.Font.Color = Color.get_Firebrick()
#set excel cell data to be italic
sheet.Range["B3:D7"].Style.Font.IsItalic = True
#Add strikethrough
sheet.Range["D3"].Style.Font.IsStrikethrough = True
sheet.Range["D7"].Style.Font.IsStrikethrough = True
#Save and Launch
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

