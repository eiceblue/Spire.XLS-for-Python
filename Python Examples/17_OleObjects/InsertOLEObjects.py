from spire.xls import *
from spire.common import *


def GenerateImage( fileName):
        book = Workbook()
        book.LoadFromFile(fileName)
        book.Worksheets[0].PageSetup.LeftMargin = 0
        book.Worksheets[0].PageSetup.RightMargin = 0
        book.Worksheets[0].PageSetup.TopMargin = 0
        book.Worksheets[0].PageSetup.BottomMargin = 0
        return book.Worksheets[0].ToImage(1, 1, 19, 5)

inputFile = "./Demos/Data/InsertOLEObjects.xls"
outputFile = "InsertOLEObjects.xlsx"

workbook = Workbook()
ws = workbook.Worksheets[0]
ws.Range["A1"].Text = "Here is an OLE Object."
#insert OLE object
image = GenerateImage(inputFile)
with Stream() as stream:
    image.Save(stream,ImageFormat.get_Png())
    oleObject = ws.OleObjects.Add(inputFile, stream, OleLinkType.Embed)
oleObject.Location = ws.Range["B4"]
oleObject.ObjectType = OleObjectType.ExcelWorksheet
#save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

