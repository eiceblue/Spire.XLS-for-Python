import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ReadFormulas.xlsx"
outputFile = "ReadFormulas.txt"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
formula = sheet.Range["C14"].Formula
value = str(sheet.Range["C14"].FormulaNumberValue)
File.AppendText(outputFile, "Formula："+formula + "\r\n"+ "Value：" + value)
workbook.Dispose()
