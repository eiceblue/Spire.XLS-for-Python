from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/Template_Xls_5.xlsx"
outputFile = "ExtractTextFromATextbox.txt"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Get the first textbox.
shape = sheet.TextBoxes[0]
#Extract text from the text box.
content = []
content.append("The text extracted from the TextBox is: ")
content.append(shape.Text)
#Save to file.
AppendAllText(outputFile, content)
workbook.Dispose()

