import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ReadImages.xlsx"
outputFile = "GetCellDisplayedText.txt"

#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Get first worksheet of the workbook
worksheet = workbook.Worksheets[0]
#Set value for B8
cell = worksheet.Range["B8"]
cell.NumberValue = 0.012345
#Set the cell style
style = cell.Style
style.NumberFormat = "0.00"
#Get the cell value
cellValue = cell.Value
#Get the displayed text of the cell
displayedText = cell.DisplayedText
#Create StringBuilder to save 
content = []
#Set string format for displaying
result = "B8 Value: " + cellValue + "\r\nB8 displayed text: " + displayedText
#Add result string to StringBuilder
content.append(result)
#Save them to a txt file
File.AppendAllText(outputFile, content)


