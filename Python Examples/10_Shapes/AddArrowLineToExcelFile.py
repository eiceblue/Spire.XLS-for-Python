from spire.xls import *
from spire.common import *


outputFile = "AddArrowLineToExcelFile.xlsx"

#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Add a Double Arrow and fill the line with solid color.
line = sheet.TypedLines.AddLine()
line.Top = 10
line.Left = 20
line.Width = 100
line.Height = 0
line.Color = Color.get_Blue()
line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow
line.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow
#Add an Arrow and fill the line with solid color.
line_1 = sheet.TypedLines.AddLine()
line_1.Top = 50
line_1.Left = 30
line_1.Width = 100
line_1.Height = 100
line_1.Color = Color.get_Red()
line_1.BeginArrowHeadStyle = ShapeArrowStyleType.LineNoArrow
line_1.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow
#Add an Elbow Arrow Connector.
line3 = sheet.TypedLines.AddLine() if isinstance(sheet.TypedLines.AddLine(), XlsLineShape) else None
line3.LineShapeType = LineShapeType.ElbowLine
line3.Width = 30
line3.Height = 50
line3.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow
line3.Top = 100
line3.Left = 50
#Add an Elbow Double-Arrow Connector.
line2 = sheet.TypedLines.AddLine() if isinstance(sheet.TypedLines.AddLine(), XlsLineShape) else None
line2.LineShapeType = LineShapeType.ElbowLine
line2.Width = 50
line2.Height = 50
line2.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow
line2.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow
line2.Left = 120
line2.Top = 100
#Add a Curved Arrow Connector.
line3 = sheet.TypedLines.AddLine() if isinstance(sheet.TypedLines.AddLine(), XlsLineShape) else None
line3.LineShapeType = LineShapeType.CurveLine
line3.Width = 30
line3.Height = 50
line3.EndArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen
line3.Top = 100
line3.Left = 200
#Add a Curved Double-Arrow Connector.
line2 = sheet.TypedLines.AddLine() if isinstance(sheet.TypedLines.AddLine(), XlsLineShape) else None
line2.LineShapeType = LineShapeType.CurveLine
line2.Width = 30
line2.Height = 50
line2.EndArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen
line2.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen
line2.Left = 250
line2.Top = 100
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

