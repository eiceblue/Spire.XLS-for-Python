from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/WAVFileSample.wav"
inputimg ="./Demos/Data/SpireXls.png"
outputFile = "InsertWavFileOLEObject.xlsx"

#Create a workbook
workbook = Workbook()
#Get the first sheet
sheet = workbook.Worksheets[0]
#Add OLE object
with Stream(inputimg) as fs:
    oleObject = sheet.OleObjects.Add(inputFile, fs, OleLinkType.Embed)
#Set the object location
oleObject.Location = sheet.Range["B4"]
#Set the object type as package
oleObject.ObjectType = OleObjectType.Package
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

