import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ProtectedWorksheet.xlsx"
outputFile = "VerifyProtectedWorksheet.txt"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
worksheet = workbook.Worksheets[0]
#Verify the first worksheet 
detect = worksheet.IsPasswordProtected
#Create StringBuilder to save 
content = []
#Set string format for displaying
result = "The first worksheet is password protected or not: " + str(detect)
#Add result string to StringBuilder
content.append(result)
#Save the document
File.AppendAllText(outputFile, content)
workbook.Dispose()
