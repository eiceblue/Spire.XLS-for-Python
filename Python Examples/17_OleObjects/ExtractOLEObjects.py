import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ExtractOle2.xlsx"
outputFile1 = "ExtractOLEObjects.docx"
outputFile2 = "ExtractOLEObjects.pdf"
outputFile3 = "ExtractOLEObjects.pptx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Extract ole objects
if sheet.HasOleObjects:
    for obj in sheet.OleObjects:
        type = obj.ObjectType
            #Word document
        if type is OleObjectType.WordDocument:
            File.WriteAllBytes(outputFile1, obj.OleData)
workbook.Dispose()


workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Extract ole objects
if sheet.HasOleObjects:
    for obj in sheet.OleObjects:
        type = obj.ObjectType
            #PDF document
        if type is OleObjectType.AdobeAcrobatDocument:
            File.WriteAllBytes(outputFile2, obj.OleData)
workbook.Dispose()


workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Extract ole objects
if sheet.HasOleObjects:
    for obj in sheet.OleObjects:
        type = obj.ObjectType
            #PowerPoint document
        if type is OleObjectType.PowerPointSlide:
            File.WriteAllBytes(outputFile3, obj.OleData)
workbook.Dispose()

