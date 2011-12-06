library(RDCOMClient)
o = COMCreate("Excel.Application")

library(SWinTypeLibs)

i = getTypeInfo(o)[[1]]	
func = getFuncs(i)[["AddChartAutoFormat"]]


#.Call("R_GetIDsOfNames", o, c("AddChartAutoFormat", "Chart", "Name", "Description"), TRUE)

# getNameIDs(func@memid, i)
# getNameIDs(func, i)