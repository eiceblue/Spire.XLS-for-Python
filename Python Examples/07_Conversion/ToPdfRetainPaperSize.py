from spire.xls import *
from spire.xls.common import *

# Create an instance of the Workbook class
book = Workbook()        

# Load an existing Excel file
book.LoadFromFile("Data/ToPDF.xlsx")        

# Set the SheetFitToPageRetainPaperSize property to True to ensure that when converting to PDF, 
book.ConverterSetting.SheetFitToPageRetainPaperSize = True     

# Save the workbook as a PDF file
book.SaveToFile("ToPdfRetainPaperSize.pdf", FileFormat.PDF)

# Dispose the workbook instance to release all resources used by this instance.
book.Dispose()