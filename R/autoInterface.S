

getConverterFunction = 
function(param)
{
 NULL
}


computeFunctionInformation =
function(desc) {

  if(length(desc@parameters) == 0)
    return(list(paraNames = NULL, required = NULL, converters = NULL))

  paramNames = tolower(names(desc@parameters))
  required = paramNames[!sapply(desc@parameters, isOptional)]
  converters = lapply(desc@parameters, getConverterFunction)
  if(all(sapply(converters, is.null))) 
   converters = NULL

  list(paramNames = paramNames,
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
   i = name %in% names(funcs)
   if(!i)
    stop("No such property or function ", name, " in the COM object of type", class(x))

  
   if(is(funcs[[name]], "PropertyGetDescription")) {
     .COM(x, name, .dispatch = 2, .ids = .nameIDs[[name]])
   } else {
    function(...) {
     args = sys.call()
     argNames = tolower(names(args)[-1])

     def = funcs[[name]]
     if(any(which <- is.na(match(argNames, def$paramNames))))
         stop("Unmatched arguments", paste(argNames[which], collapse=", "))

     if(any(is.na(match(def$required, argNames))))
       stop("Required argument(s) missing:", paste(def$required[is.na(match(def$required, argNames))], collapse=", "))

     .COM(x, name, ..., .dispatch = 1, .ids = .nameIDs[[name]])
    }
   }
  }
  setMethod("$", className,  Get)

  if(length(which) > 0 && all(which))
    .SetProperties <- funcs[!which]
  else
    .SetProperties <- list()

  Set <- function(x, name, value) {
     i = tolower(name) %in% names(.SetProperties)
     if(!i)
        stop("No mechainsm for setting the property ", name)

     .COM(x, name, value, .dispatch = 4)
  }

  setMethod("$<-", className, Set)

  invisible(list(Get, Set))
}




createCOMSClass =
function(libEntry, className, where = NULL)
{
  setClass(className, "COMIDispatch")
  setAs("COMIDispatch", className, function(from) {
    tmp = new("Workbooks")
    tmp@ref = from@ref
    tmp
  })

  generateOperators(libEntry, className)

  if(!exists(".COMSClassMap"))
    stop("Create .COMSClassMap")

  .COMSClassMap[[libEntry@guid]] <<- className

  className
}

