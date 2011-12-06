library(RDCOMClient)
library(SWinTypeLibs)
o = COMCreate("Excel.Application")
e = new("ExcelApplication")
e@ref = o@ref

# e$ActiveWindow
