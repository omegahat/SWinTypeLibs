library(SWinTypeLibs)
l = LoadTypeLib("c:\\WINNT\\system32\\wmpui.dll")
getTypeDocumentation(l, 14)
getInfoTypes(l)
