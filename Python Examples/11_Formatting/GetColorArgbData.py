import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *

inputFile = "./Demos/Data/templateAz.xlsx"
outputFile = "GetColorArgbData.txt"

#Create a workbook
workbook = Workbook()
#Load a excel document
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
strB = []
#Get font color
color1 = sheet.Range["B2"].Style.Font.Color
#Read ARGB data of Color
strB.append("The font color of B2: ARGB=({0},{1},{2},{3})".format(color1.A, color1.R, color1.G, color1.B))
color2 = sheet.Range["B3"].Style.Font.Color
strB.append("The font color of B3: ARGB=({0},{1},{2},{3})".format(color2.A, color2.R, color2.G, color2.B))
color3 = sheet.Range["B4"].Style.Font.Color
strB.append("The font color of B4: ARGB=({0},{1},{2},{3})".format(color3.A, color3.R, color3.G, color3.B))
#Save to file
File.AppendAllText(outputFile, strB)
workbook.Dispose()
