from spire.xls import *
from spire.xls.common import *
        
# Create a workbook instance
workbook = Workbook()

# Load an existing Excel document
workbook.LoadFromFile("Data/ImageHyperlink.xlsx")

# Get the first worksheet in the workbook
sheet = workbook.Worksheets.get_Item(0)

# Get the first picture found in the worksheet
picture = sheet.Pictures.get_Item(0)

# Retrieve the hyperlink attached to the picture
link = picture.GetHyperLink()

# Extract the address (URL or file path) of the hyperlink
address = link.Address

# Print the extracted address to the console
print(address)