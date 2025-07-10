from spire.xls import *
from spire.xls.common import *

def AppendText(fname:str,text:str):
    fp = open(fname,"w")
    fp.write(text + "\n")
    fp.close()

inputFile = "./Demos/Data/ReadFormulas.xlsx"
outputFile = "ReadFormulas.txt"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
formula = sheet.Range["C14"].Formula
value = str(sheet.Range["C14"].FormulaNumberValue)
AppendText(outputFile, "Formula："+formula + "\r\n"+ "Value：" + value)
workbook.Dispose()
