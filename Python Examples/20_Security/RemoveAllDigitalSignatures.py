from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WithDigitalSignature.xlsx"
outputFile = "RemoveAllDigitalSignatures.xlsx"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Remove all digital signatures.
workbook.RemoveAllDigitalSignatures()
#Save to file.
workbook.SaveToFile(outputFile, FileFormat.Version2013)
workbook.Dispose()

