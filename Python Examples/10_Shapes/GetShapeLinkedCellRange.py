from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/CellLinkedRangeLocal.xlsx"
outputFile = "GetShapeLinkedCellRange.xlsx"

#create a workbook
workbook = Workbook()

#load an excel document
workbook.LoadFromFile(inputFile)

#get the first worksheet
sheet=workbook.Worksheets[0]

sb = []

#get PrstGeomShapes from sheet
prstGeomShapeCollection = sheet.PrstGeomShapes

#get shape
shape = prstGeomShapeCollection["Yesterday"]

#get shape linked cell range
cellAddress = shape.LinkedCell.RangeAddress

#append in sb
sb.append(cellAddress + "\n")

shape = prstGeomShapeCollection["NewShapes"]
cellAddress = shape.LinkedCell.RangeAddress
sb.append(cellAddress)

#save to txt file
AppendAllText(outputFile, sb)
workbook.Dispose()


