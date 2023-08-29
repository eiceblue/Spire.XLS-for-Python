import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/EncryptedFile.xlsx"
outputFile = "OpenEncryptedFile.txt"

#Create string builder
builder = []
passwords = ["password1", "password2", "password3", "1234"]
i = 0
while i < len(passwords):
    try:
        #Create a workbook
        workbook = Workbook()
        #Open password
        workbook.OpenPassword = passwords[i]
        #Load the document
        workbook.LoadFromFile(inputFile)
        builder.append("Password = " + passwords[i] + " is correct." + " The encrypted Excel file opened successfully!")
    except ArgumentError as ex:
        builder.append("Password = " + passwords[i] + "  is not correct")
    i += 1
#Save to txt file
File.AppendAllText(outputFile, builder)


