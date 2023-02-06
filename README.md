# AutomaticPDFPrinter
This program takes a list of machine part numbers from an Excel file (.xlsx or .csv), and searches for a PDF drawing of the part thorugh an internal file directory. The part numbers are often 7 digits long; however, some part numbers contained dashes or were missing leading zeros, which did not match up with the PDFs in the drawing directory. The program simplifies these part numbers if they cannot be found.

The user can update the starting row and column of their list, which sheet to print from, as well as choose a folder directory to create an output file called "MissingParts.txt", which contains the part numbers of all PDFs that could not be found.

All files are automatically printed through PowerShell to the user's default printer. PowerShell was utilized because it is an interpreted language, and the machinery company for which this program was made had strict security regulations surrounding compiled languages. 

See the ExampleDrawings branch for a collection of test PDFs.
