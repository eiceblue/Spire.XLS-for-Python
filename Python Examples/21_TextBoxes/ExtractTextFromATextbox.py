import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


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
File.AppendAllText(outputFile, content)
workbook.Dispose()

