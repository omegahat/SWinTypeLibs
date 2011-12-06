library(SWinTypeLibs)
lib = LoadTypeLib("C:/Program Files/Common Files/system/ado/msado15.dll")
getTypeLibTypes(lib)

names(lib)

cmd = lib[["Command"]]

k = .Call("R_getCoClass", cmd)[[1]]

getFuncs(k)


