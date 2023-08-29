from spire.xls import *
from spire.common import *


outputFile = "WrapOrUnwrapTextInCells.xlsx"

#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Wrap the excel text
sheet.Range["C1"].Text = "e-iceblue is in facebook and welcome to like us"
sheet.Range["C1"].Style.WrapText = True
sheet.Range["D1"].Text = "e-iceblue is in twitter and welcome to follow us"
sheet.Range["D1"].Style.WrapText = True
#Unwrap the excel text
sheet.Range["C2"].Text = "http://www.facebook.com/pages/e-iceblue/139657096082266"
sheet.Range["C2"].Style.WrapText = False
sheet.Range["D2"].Text = "https://twitter.com/eiceblue"
sheet.Range["D2"].Style.WrapText = False
#Set the text color of Range["C1:D1"]
sheet.Range["C1:D1"].Style.Font.Size = 15
sheet.Range["C1:D1"].Style.Font.Color = Color.get_Blue()
#Set the text color of Range["C2:D2"]
sheet.Range["C2:D2"].Style.Font.Size = 15
sheet.Range["C2:D2"].Style.Font.Color = Color.get_DeepSkyBlue()
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

