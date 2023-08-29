from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/SpireXls.png"
outputFile = "AddImageHyperlink.xlsx"

#Create a workbook
workbook = Workbook()
sheet = workbook.Worksheets[0]
#Add the description text
sheet.Columns[0].ColumnWidth = 22
sheet.Range["A1"].Text = "Image Hyperlink"
sheet.Range["A1"].Style.VerticalAlignment = VerticalAlignType.Top
#Insert an image to a specific cell
picture = sheet.Pictures.Add(2, 1, inputFile)
#Add a hyperlink to the image
picture.SetHyperLink("https://www.e-iceblue.com/Introduce/excel-for-net-introduce.html", True)
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

