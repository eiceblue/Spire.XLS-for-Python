from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

def Exists(pname:str):
    exists = os.path.exists(pname)
    return exists

def CreateDirectory(pname:str):
    if os.path.exists(pname) :
        return
    os.makedirs(pname)

inputFile = "./Demos/Data/Template_Xls_5.xlsx"
outputFile = "ExtractTextImageFromShape.txt"
outputFile_i = "Output/Image/"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Extract text from the first shape and save to a txt file.
shape1 = sheet.PrstGeomShapes[2]
s = shape1.Text
sb = []
sb.append("The text in the third shape is: " + s)
AppendAllText(outputFile, sb)
workbook.Dispose()


if Exists(outputFile_i) == False:
    CreateDirectory(outputFile_i)
#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Extract image from the second shape and save to a local folder.
shape2 = sheet.PrstGeomShapes[1]
image = shape2.Fill.Picture
filename = outputFile_i + "ExtractTextImageFromShape.png"
image.Save(filename, ImageFormat.get_Png())
image.Dispose()
workbook.Dispose()
