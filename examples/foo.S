
lib = LoadTypeLib("C:\\Program Files\\Microsoft Office\\Office\\EXCEL9.OLB")

els = getTypeInfo(lib)

enums = els[sapply(els, function(x) names(x@type) == "enum")]

setClass("EnumerationDefinition",
           representation("integer"))

setAs("integer", "EnumerationDefinition",
        function(from) {
          if(length(names(from)) == 0)
           stop("Enumeration must have named elements")
          x = new("EnumerationDefinition")
          x@.Data = from
          x
        })

setClass("EnumerationValue", representation("integer"))

setClass("XlCellnew", representation("EnumerationDefinition")
setClass("XlCellType", representation("EnumerationValue"))
setAs("


workbooks = lib[["Workbooks"]]
funcs = getFuncs(workbooks)

 names(funcs)

f = funcs[["Open"]]
names(f@parameters)

 sapply(f@parameters, isOptional)

  obj$Open(filename, ...)


$.Workbook =
 function(x, name) {
   i = name %in% names(propertyGet)
   if(!is.na(i)) {
       # Use the propertyGet[[i]] to control the conversion of the return type.
      .COM(x, name, .dispatch = 2)
  } else if(!is.na(i <- name %in% names(functions))) {
     function(...) {
      .COM(x, name, ..., .dispatch = 1)     
     }
  } else 
    stop("No such property or function in the COM object of type", class(x)
   
 }


  wkbs$Open("foo.xls")


function(x, name) {

 .paramNames = tolower(names(desc@paramters))
 .required = .paramNames[!sapply(desc@parameters, isOptional)]
 .converters = lapply(desc@parameters, getConverterFunctions)
 if(all(sapply(.converters, is.null))) 
   .converters = NULL

 function(...) {
    args = sys.call()
    argNames = tolower(names(args)[-1])
    if(any(which = is.na(match(argNames, .paramNames))))
         stop("Unmatched arguments", paste(argNames[which], collapse=", "))

    if(any(is.na(match(.required, argNames))))
       stop("Required argument(s) missing:", paste(.required[is.na(match(.required, argNames))], collapse=", "))
    .COM(x, name, ..., .dispatch)
 }
}


 getNameIDs(func, wks)

 getNameIDs(wks)


function(x, name) {

  i = name %in% names(propertyGet)
  if(!is.na(i)) {
      .COM(x, name, .dispatch = 2, .ids = nameIDs[[name]])
  else if(!is.na(i <- name %in% names(functions))) {
    function(...)
     .COM(x, name, ..., .dispatch, .ids = nameIDs[[name]])
  } else
    stop(...)
}


  funcs = getFuncs(libEntry)
  which = sapply(funcs, function(x) !is(x, "PropertySetDescription"))
  nameIDs = getNameIDs(libEntry)[which]

computeFunctionInformation =
function(desc) {

  paramNames = tolower(names(desc@paramters))
  required = .paramNames[!sapply(desc@parameters, isOptional)]
  converters = lapply(desc@parameters, getConverterFunctions)
  if(all(sapply(converters, is.null))) 
   converters = NULL

  list(paraNames = paramNames,
       required = required, 
       converters = converters)
}

generateOperators =
function(libEntry, className) {

  funcs = getFuncs(libEntry)

  which = sapply(funcs, function(x) !is(x, "PropertySetDescription"))
  .nameIDs = getNameIDs(libEntry)[which]

  getDefs = lapply(funcs[which], computeFunctionInformation)

  Get = 
   function(x, name) {
   i = name %in% names(propertyGet)
   if(is.na(i))
    stop("No such property or function ", name, " in the COM object of type", class(x))

  
   if(is(funcs[[name]], "PropertyGetDescription")) {
     .COM(x, name, .dispatch = 2, .ids = .nameIDs[[name]])
   } else {
    function(...) {
     args = sys.call()
     argNames = tolower(names(args)[-1])

     def = funcs[[name]]
     if(any(which = is.na(match(argNames, def$paramNames))))
         stop("Unmatched arguments", paste(argNames[which], collapse=", "))

     if(any(is.na(match(def$required, argNames))))
       stop("Required argument(s) missing:", paste(def$required[is.na(match(def$required, argNames))], collapse=", "))

     .COM(x, name, ..., .dispatch = 1, .ids = .nameIDs[[name]])
    }
   }
  }
  setMethod("$", className,  Get)

  invisible(Get)
}

processFunctions =
function(obj)
{

 funcs = getFuncs(obj)
 readProps = names(funcs)[sapply(funcs, function(x) x@invokeType == 2)]
 writeProps = names(funcs)[sapply(funcs, function(x) x@invokeType == 4)]

 readWriteProps = readProps[!is.na(match(readProps, writeProps))]

 readOnlyProps = readProps[is.na(match(readProps, writeProps))] 
 writeOnlyProps = writeProps[is.na(match(writeProps, readProps))] 


callable  = names(funcs)[sapply(funcs, function(x) x@invokeType == 1)]

list(readOnly = readOnlyProps, writeOnly = writeOnlyProps,
     readWrite = readWriteProps,  functions = callable)
}


comAccessor =
function(info, className, where = NULL) {

  get <- function(x, name) {
    name <- tolower(name)

    if(!is.na(match(name, info$functions))) {
       return(function(...) {
                .COM(x, name, ...)
              })
    } else if(!is.na(match(name, c(info$readOnly, info$readWrite)))) {
       .Call("R_getProperty",x, as.character(name), NULL)
    } else {
       browser()
       stop("No such element ", name, " in COM object")
    }
  }

  set <- function(x, name, value) {
    name <- tolower(name)

    if(!is.na(match(name, c(info$readWrite, info$writeOnly)))) {
       .Call("R_setProperty", x, as.character(name), list(value))
       x
    } else
      stop("No writeable property named ", name, " in COM object")
  }

  setMethod("$", className, get, where = where)
  setMethod("$<-", className, set, where = where)
}

