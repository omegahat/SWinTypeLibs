library(RDCOMClient)
o = COMCreate("Excel.Application")
library(SWinTypeLibs)
ff = getFuncs(o)
f = ff[["Worksheets"]]

