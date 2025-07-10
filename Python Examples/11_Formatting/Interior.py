from spire.xls import *
from spire.xls.common import *
import random
import math


outputFile = "Interior.xlsx"

#Create a workbook
workbook = Workbook()
#Initialize the workbook
sheet = workbook.Worksheets[0]
#Specify the version
workbook.Version = ExcelVersion.Version2007
#Define the number of the colors
maxColor = len(ExcelColors)

#Create a random object
random.seed(10000000)
for i in range(2, 40):
    #Random backKnownColor
    backKnownColor = ExcelColors(random.randint(1, math.trunc(maxColor / float(2))))
    #Add text
    sheet.Range["A1"].Text = "Color Name"
    sheet.Range["B1"].Text = "Red"
    sheet.Range["C1"].Text = "Green"
    sheet.Range["D1"].Text = "Blue"
    #Merge the sheet"E1-K1"
    sheet.Range["E1:K1"].Merge()
    sheet.Range["E1:K1"].Text = "Gradient"
    sheet.Range["A1:K1"].Style.Font.IsBold = True
    sheet.Range["A1:K1"].Style.Font.Size = 11
    #Set the text of color in sheetA-sheetD
    colorName = str(backKnownColor)
    sheet.Range["A{0}".format(i)].Text = colorName
    sheet.Range["B{0}".format(i)].NumberValue = workbook.GetPaletteColor(backKnownColor).R
    sheet.Range["C{0}".format(i)].NumberValue = workbook.GetPaletteColor(backKnownColor).G
    sheet.Range["D{0}".format(i)].NumberValue = workbook.GetPaletteColor(backKnownColor).B
    #Merge the sheets
    sheet.Range["E{0}:K{0}".format(i)].Merge()
    #Set the text of sheetE-sheetK
    sheet.Range["E{0}:K{0}".format(i)].Text = colorName
    #Set the interior of the color
    sheet.Range["E{0}:K{0}".format(i)].Style.Interior.FillPattern = ExcelPatternType.Gradient
    sheet.Range["E{0}:K{0}".format(i)].Style.Interior.Gradient.BackKnownColor = backKnownColor
    sheet.Range["E{0}:K{0}".format(i)].Style.Interior.Gradient.ForeKnownColor = ExcelColors.White
    sheet.Range["E{0}:K{0}".format(i)].Style.Interior.Gradient.GradientStyle = GradientStyleType.Vertical
    sheet.Range["E{0}:K{0}".format(i)].Style.Interior.Gradient.GradientVariant = GradientVariantsType.ShadingVariants1
#AutoFit Column
sheet.AutoFitColumn(1)
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()


