from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/ReadComment.xls"
outputFile = "ReadComment.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
builder = []
builder.append(sheet.Range["A1"].Comment.Text+"\n\t")
builder.append(str(sheet.Range["A2"].Comment.RichText.RtfText))
AppendAllText(outputFile, builder)
workbook.Dispose()



