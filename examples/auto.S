process =
function(obj)
{
 funcs = getFuncs(obj)
 readProps = names(funcs)[sapply(funcs, function(x) x@invokeType == 2)]
 writeProps = names(funcs)[sapply(funcs, function(x) x@invokeType == 4)]

 readWriteProps = readProps[!is.na(match(readProps, writeProps))]

 readOnlyProps = readProps[is.na(match(readProps, writeProps))] 
 writeOnlyProps = writeProps[is.na(match(writeProps, readProps))] 


 callable  = names(funcs)[sapply(funcs, function(x) x@invokeType == 1)]

 names(funcs) <- tolower(names(funcs))
 funcs = funcs[!is.na(match(names(funcs),  tolower(callable)))]

list(readOnly = tolower(readOnlyProps), writeOnly = tolower(writeOnlyProps),
     readWrite = tolower(readWriteProps),  functions = tolower(callable),
     functionDefinitions = funcs)
}