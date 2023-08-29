import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


outputFile = "GetExcelPaperDimensions.txt"

#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
content = []
#Get the dimensions of A2 paper.
sheet.PageSetup.PaperSize = PaperSizeType.A2Paper
content.append("A2Paper: " + str(sheet.PageSetup.PageWidth) + " x " + str(sheet.PageSetup.PageHeight))
#Get the dimensions of A3 paper.
sheet.PageSetup.PaperSize = PaperSizeType.PaperA3
content.append("PaperA3: " + str(sheet.PageSetup.PageWidth) + " x " + str(sheet.PageSetup.PageHeight))
#Get the dimensions of A4 paper.
sheet.PageSetup.PaperSize = PaperSizeType.PaperA4
content.append("PaperA4: " + str(sheet.PageSetup.PageWidth) + " x " + str(sheet.PageSetup.PageHeight))
#Get the dimensions of paper letter.
sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter
content.append("PaperLetter: " + str(sheet.PageSetup.PageWidth) + " x " + str(sheet.PageSetup.PageHeight))
#Save to file
File.AppendAllText(outputFile, content)
workbook.Dispose()

