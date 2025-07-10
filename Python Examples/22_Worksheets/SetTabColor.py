from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SetTabColor.xlsx"
outputFile = "SetTabColor.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Set the tab color of first sheet to be red 
worksheet = workbook.Worksheets[0]
worksheet.TabColor = Color.get_Red()
#Set the tab color of first sheet to be green 
worksheet = workbook.Worksheets[1]
worksheet.TabColor = Color.get_Green()
#Set the tab color of first sheet to be blue 
worksheet = workbook.Worksheets[2]
worksheet.TabColor = Color.get_LightBlue()
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

