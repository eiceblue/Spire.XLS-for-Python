from spire.xls import *
from spire.xls.common import *

def WriteAllBytes(fname:str,data):
    fp = open(fname,"wb")
    for d in data:
        fp.write(d)
    fp.close()

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
            WriteAllBytes(outputFile1, obj.OleData)
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
            WriteAllBytes(outputFile2, obj.OleData)
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
            WriteAllBytes(outputFile3, obj.OleData)
workbook.Dispose()

