library(RDCOMClient)
library(SWinTypeLibs)


lib = LoadTypeLib("C:/Microsoft Office/OFFICE11/Excel.exe")

e = COMCreate("Excel.Application")
l = getTypeLib(e)
info = getTypeLibInfo(l)

lreg = .Call("R_loadRegTypeLib", info$guid,  
                c(info$major, info$minor), NULL, PACKAGE = "SWinTypeLibs")
