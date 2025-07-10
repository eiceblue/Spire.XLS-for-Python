from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()
	

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
    except SpireException as ex:
        builder.append("Password = " + passwords[i] + "  is not correct")
    i += 1
#Save to txt file
AppendAllText(outputFile, builder)


