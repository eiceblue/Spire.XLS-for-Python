import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "GetCellValueByCellName.txt"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Specify a cell by its name.
cell = sheet.Range["A2"]
content = []
#Get vaule of cell "A2".
content.append("The vaule of cell A2 is: " + cell.Value) #Save to file.
File.AppendAllText(outputFile, content)

