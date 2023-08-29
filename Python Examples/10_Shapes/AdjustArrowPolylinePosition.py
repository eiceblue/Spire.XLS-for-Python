from spire.xls import *
from spire.common import *


outputFile = "AdjustArrowPolylinePosition.xlsx"

workbook = Workbook()
worksheet = workbook.Worksheets[0]
#Draw an elbow arrow
line = worksheet.TypedLines.AddLine(5, 5, 100, 100, LineShapeType.ElbowLine)
line.EndArrowHeadStyle = ShapeArrowStyleType.LineNoArrow
line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow
ad = line.ShapeAdjustValues.AddAdjustValue(GeomertyAdjustValueFormulaType.LiteralValue)
#When the parameter value is less than 0, the focus of the line is on the left side of the left point, when it is equal to 0, the position is the same as the left point, it is equal to 50 in the middle of the graph, and when it is equal to 100, it is the same as the right point.
ad.SetFormulaParameter([-50])
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

