library(RDCOMClient)
library(SWinTypeLibs)

e = COMCreate("Excel.Application")

f0 = getFuncs(e)


lib = LoadTypeLib(e)

names(lib)
app = lib[["Application"]]



# LoadTypeLib() was failing so narrowed it down to here.
# Problem in getTypeInfo
# i = .Call("R_getDCOMInfoEntry", e@ref, 0L, PACKAGE = "SWinTypeLibs")


f2 = getFuncs(lib[["Workbooks"]])

# Also
wbooks = e[["Workbooks"]]
f = getFuncs(wbooks)


# This gives NULL
#funs = getFuncs(lib[["Worksheet"]])