library(SWinTypeLibs)
library(RDCOMClient)

e = COMCreate("Excel.Application")
info = getTypeInfo(e, 1)
tmp = getFuncs(info)
 # should be the same thing
tmp1 = getFuncs(e)

 # Now get the details for the Workbooks class.
books = e[["Workbooks"]]
binfo = getFuncs(books)
