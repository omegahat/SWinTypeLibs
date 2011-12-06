library(RDCOMClient)
e = COMCreate("Excel.Application")

library(SWinTypeLibs)

i = .Call("R_getRefTypeInfo", getTypeInfo(e)[[1]], as.integer(-1))

