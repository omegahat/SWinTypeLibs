library(SWinTypeLibs)
lib = LoadTypeLib("C:\\Program Files\\Microsoft Office\\Office\\EXCEL9.OLB")

wks = lib[["Worksheet"]]	
i = getElements(wks)[[1]]
funcs = getFuncs(i)
esel = funcs[["EnableSelection"]]
ii = getTypeInfo(i, esel@returnType)
def = getElements(ii)


wks = lib[["Workbooks"]]	
func = getFuncs(wks)[["Open"]]
ii = getTypeInfo(i, func@returnType)
def = getElements(ii)
