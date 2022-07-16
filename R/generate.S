# Enumerations and finish off the events.
# Do we have an enumeration class? Borrow it from RGtk or RSWIG.

INamedTypeLib = 
function(lib)
{
#XXX This ends up calling the C routine for getting the names (getTypeLibNames) 3 times.
  n = new("INamedTypeLib", lib)
  n@elNames = names(lib) 
  n@uuids = getTypeLibUUIDs(lib)
  n@infoEls = getTypeInfo(lib, 1:length(lib))

  n
}


# Define a generic class for the functions we will create 
# to access COM properties or methods.
# The elements of the vector represent distinctc _lines_ of code 

setClass("COMOperationDefinition", 
          representation("character",
                         referencedClasses = "list"
	                 ))

# A way to turn the code into a string.
# call with as(x, "character"), rather than as.character()

setAs("COMOperationDefinition", "character",
       function(from) {
         paste(c("", rep("\t", length(from)-2), ""), from, sep="", collapse = "\n")
       })

# Method for providing a convenient display on the terminal/R output
# of a COMOperationDefinition.
setMethod("show", "COMOperationDefinition",
           function(object) {
             cat(as(object, "character"))
           })


# A specific class derived from COMOperationDefinition and this is 
# for representing functions to call COM methods.
# The idea is that we can use this to maintain the definition
# as a character vector of lines, but to look at it easily as
# a single string.  The lines of code can make it easier to
# indent the entire code and use it in different contexts,
# i.e. levels of indentation.
setClass("COMMethodDefinition", contains = "COMOperationDefinition")
setClass("COMEventDefinition", contains = "COMMethodDefinition")
setClass("COMPropertySetDefinition", contains = "COMOperationDefinition")
setClass("COMPropertyGetDefinition", contains = "COMOperationDefinition")


# A class for describing the generated code for all of the elements
# in a Type Library.  Essentially this is a list of RCOMTypeInterfaceDefinition
# elements (defined below), and we are giving it a type so we can create
# methods such as define(), writeCode(), etc. for it.
setClass("RCOMTypeLibraryInterfaceDefinition",
           representation(types = "list",
                          enums = "list",
                          events = "list",
                          typeEventMap = "list",
                          uuids = "character",
                          coclass = "character", # a named character vector
                          referencedClasses = "list",
	                  externalClasses = "list")
        )

generateEnums =
function(els, lib, verbose = FALSE)
{
  lapply(els, function(el) {
                 getElements(el)
              })
}

generateInterface =
#
# Generate code and the names of classes from a type library
#
#
function(lib, classes = names(lib), events = TRUE,
           enums = TRUE,
	   coclass = TRUE,
           defaultClass = "CompiledCOMIDispatch",
           defaultListClass = c("COMList", "COMTypedList"),
           class = "RCOMTypeLibraryInterfaceDefinition",
	   verbose = FALSE,
           processExternalClasses = TRUE,
           computeReferencedClasses = TRUE)
{
  if(!is(lib, "INamedTypeLib")) {
     if(verbose)
        cat("Reading library\n")
     lib = INamedTypeLib(lib)
  }

  if(is.character(class))
     obj = new(class)
  else 
     obj = class

  obj@uuids = getTypeLibUUIDs(lib)
  
   # Get just the elements we want from the type library.
  els = getElements(lib)[classes]
  if(verbose)
       cat("got", length(els), "elements in the library\n")

  if(length(els)) {
     dispatch = els[sapply(els, is, "ITypeInfoDispatch")]
     if(verbose)
       cat("Generating class information for IDispatch types")

     # Iterate over each of the elements
     #XXX Need to get the name of the new class.
     functions = lapply(names(dispatch), generateClass, lib, defaultClass, defaultListClass, verbose = verbose)
     names(functions) = names(dispatch)

     obj@types = functions
  }

  if(is.logical(enums) && enums)
     enums = lib[sapply(lib, is, "ITypeInfoEnum")]

  if(length(enums)) {
     if(is.character(enums))
         enums = lib[enums]

   if(verbose)
       cat("Processing enumeration definitions\n")

     enums = generateEnums(enums, lib)
     obj@enums = enums
  }


  if(is.logical(coclass) && coclass) {
     coclass = names(lib[sapply(lib, is, "ITypeInfoCoClass")])
  } else if(is.list(coclass)) {
     coclass = names(coclass)
  }

  if(length(coclass))  {
     coclass = structure(paste("_", coclass, sep = ""), names = coclass)

     i = match(coclass, names(lib))
     obj@coclass = structure(coclass[!is.na(i)], names = names(coclass)[!is.na(i)]) 
   } else {
     obj@coclass = coclass
   }

   # Allow for a logical value or a character vector giving the names of the 
   # events. 
  if(is.logical(events) && events)
     events = getTypeLibEventElements(lib)

  if(length(events)) {
     if(is.character(events))
        events = lib[events]

       # old way dealing directly with event interface objects.
#     obj@events = lapply(names(events), function(x) {
#                                   generateEventInterface(events[[x]], x, lib)
#                                 })
#    names(obj@events) = names(events)     
       # this way deals with the event interfaces grouped by the (co)classes
       # that support them.  There is a possibility of duplication
       # which we could avoid by generating the interface via an  unlist
       # and then regrouping.

     obj@events = lapply(events, function(x) {
                                   tt = lapply(names(x), function(id)
                                                          generateEventInterface(x[[id]], id, lib))
                                   names(tt) = names(x)
                                   tt
                                 })     

   
     obj@typeEventMap =  computeTypeEventMap(lib)
  }

  if(processExternalClasses) {
    externalClasses = getExternalClassDefs(obj)
    if(length(externalClasses)) 
       obj@externalClasses = generateExternalClassInterfaces(externalClasses)
  }

  if(computeReferencedClasses)  {
    defaultClass = character()
    obj@referencedClasses = getReferencedClasses(obj, lib, defaultClass, defaultListClass)
  }


  obj
}

computeTypeEventMap =
      # compute the association with the event interfaces
      # and the classes in the library that implement them
      # which then allows us to take an object and find what event
      # interfaces it provides, e.g. Document

      # returns a list.
function(lib)
{
   is.coclass = sapply(lib, is,  "ITypeInfoCoClass")
   coclassTypes = lapply(lib[is.coclass], function(x) grep("Events", names(getElements(x)), value = TRUE))
    coclassTypes[ sapply(coclassTypes, length) > 0 ]
}

getTypeLibEventElements =
#
# This is a function to find all the elements in the
# type library (lib) that are event interfaces.
# There are many possible heuristics for doing this. I don't know
# if there is a simple way to get at this
#
function(lib)
{
   els = getElements(lib)
   co = els[sapply(els, is, "ITypeInfoCoClass")]
   els = lapply(co, function(x) {
                 x = getElements(x)
                 i = grep("Events", names(x))
                 if(length(i))
                   x[i]
                 else
                   NULL
              })

   els[!sapply(els, is.null)]
 }

ogetTypeLibEventElements =
# This is an alternative that seems to work.
# We can compare all the I* to * and check that they have
# the same methods (less the 7 from the IDispatch & IUnknown) and 
# also have the same method ids.
#
#
function(lib)
{
   idx = grep("Events", names(lib), value = TRUE)
   if(length(idx) == 0)
    return(list())

   idx = grep("^[^I]", idx, value = TRUE)
# Check that these are in the list without the I also.
   els = lib[idx]

   els
}




#XXX Are these used? Not in this file.
if(FALSE) {
   setGeneric("getClassDefs", function(x) standardGeneric("getClassDefs"))

   setMethod("getClassDefs",  "RCOMTypeLibraryInterfaceDefinition",
              function(x) {
                sapply(x@types, function(x) x@className)
              })
}


setGeneric("writeCode", 
             function(def, file = stdout(), context = NULL, 
	                prepend = character(),
# paste(system.file("runtime", package = "SWinTypeLibs"), c("common.S"), sep = .Platform$file.sep), 
                        enumsAsValues = TRUE, 
                        namespace = "NAMESPACE", ...)
               standardGeneric("writeCode"))



writeClasses =
# output the collection of class definitions from a RCOMTypeLibraryInterfaceDefinition
# to a connection. This  allows them all to be defined in one place early in the
# code.
function(def, con)
{
  classNames =  sapply(def@types, 
                         function(x) {
                            cat(classDefinitionString(x@className), "\n", file = con, sep = "")   
                            x@className@className
                         })


   if(length(def@referencedClasses)) {
    classNames = c(classNames, 
                    sapply(def@referencedClasses, 
                      function(classDef) { 
                           # should we "compile"/process these too.
                           #XXX fix this.
                         cat(classDefinitionString(classDef), "\n", file = con, sep = "")
                         classDef@className
                     }))
   }

   if(length(def@coclass)) {
           # for the namespace.
       w = def@coclass %in% classNames
       classNames = c(classNames, names(def@coclass)[w])
   }

  list(classes = classNames, variable = character(), methods = character())
}


writeEnums =
#
#
# Write code for all the enumerations.
#
function(def, con, asValues = FALSE)
{
  if(length(def@enums) == 0)
    return(FALSE)

  if(asValues)
    vars = structure(vector("list", length(def@enums)), names = names(def@enums))
  else
    vars = character()

  classDefs = sapply(names(def@enums), 
                       function(id) {
                          writeEnumCode(def@enums[[id]], id, con, asValues  = asValues)
                          if(asValues)
                            vars[[id]] <<- names(def@enums[[id]])
                       })

  if(asValues)
    vars = unlist(vars)

  vars = c(vars, paste(names(classDefs), "Enum", sep = ""))

  list(variables = vars, classes = names(classDefs))
}

getEnumClassName = 
function(name)
{
# paste(name, "EnumValue", sep = "")
  name
}

writeEnumCode =
#
# Write code for an individual enumeration.
#
function(def, name, con, asValues = FALSE)
{
  # Define a class for this and provide the setAs() methods.
  classes = getEnumClassName(name)
  cat('\n\nsetClass("', classes, '", contains = "EnumValue")\n', file = con, sep = "")
  for(i in c("integer", "numeric", "character")) {
          cat('setAs("', 
	        i, 
	        '", "', 
	        classes, 
	        '", function(from) EnumValue(from, obj = new("', 
	        classes, 
	        '")))\n', 
	        file = con, sep = "")
  }


  # new("NamedClassDefinition", 
  #                   code = c(classes, "EnumValue"),
  #                   className = classes)

   if(asValues) {
     sapply(names(def), 
             function(id)  {
               cat( backquote(id), " = EnumValue('",  id, "', ", def[[id]], ", new('", classes , "'))\n", file = con, sep = "")
           })
     names = names(def)
   } else {
      names = character()
   }

   name = paste(classes, "Enum", sep = "")
   cat("'", name, "' = c(\n",  sep = "", file = con)

   tmp = paste(paste("   '", names(def), "'", sep = ""),
               def, sep = " = ", collapse = ",\n")   
   cat(tmp, "\n", ")\n", file = con)

   cat("storage.mode(`", name, "` ) = 'integer'\n", sep = "", file = con)

#   names = c(names, name)

   classes
}

backquote =
function(x)
  paste('`', x, '`', sep = "")	
	
writeEventCode =
function(eventDefs, con)
{
  funcs = sapply(eventDefs, function(x) { 
                              cat(x@constructor, "\n", file = con)
                              x@constructorName
                            })
  list(variables = funcs, classes = list())
}


# If we have a directory, then write individual files to that directory
# Otherwise, if we are given a filename, we write to that.
setMethod("writeCode", c("RCOMTypeLibraryInterfaceDefinition", "character"),
           function(def, file = stdout(), context = NULL,
	  	     prepend = character(),
                      enumsAsValues = TRUE,
                      namespace = "NAMESPACE",
                      ...) {

              ns = list()

              isDir = file.info(file)$isdir			  
              if(!is.na(isDir) && isDir) {
	         # We were given a directory, so create separate
	         # files in for the different classes.
	
                 #XXX copy the prepend files to this directory.


                 classes = paste(file, "classes.S", sep = .Platform$file.sep)
	         con = file(classes)
  	         open(con, "w")
                 ns$Classes = writeClasses(def, con)
                 close(con)

                 if(length(def@enums)) {
                    classes = paste(file, "enums.S", sep = .Platform$file.sep)
   	            con = file(classes)
  	            open(con, "w")
                    ns$Enums = writeEnums(def, con, enumsAsValues)
                    close(con)
                 }

	         for(i in prepend) 
                     file.copy(i, file)


 	 	 if(FALSE && length(def@events)) {
                    EventTableFile = paste(file, "EventTable.rda", sep = .Platform$file.sep)                   
                    classes = paste(file, "EventConstructors.S", sep = .Platform$file.sep)
    	            con = file(classes)
	            open(con, "w")
                    ns$Events = writeEventCode(def@events, con)
                    close(con)
          	 }

# Need to know the order in which to source these files.

               sapply(def@types,
                      function(x) {
                        f = paste(file, paste(x@className@className[1], "S", sep = "."), sep = .Platform$file.sep)
                        writeCode(x, file = f, context, ...)
                     })

                 if(!is.na(namespace) && length(namespace) > 0)  
      		     namespace = paste(file, namespace, sep = .Platform$file.sep)

              } else {
   	              # Writing to a single file.
	        if(length(def@externalClasses))
                   def =  mergeInterfaces(def)
#   	           def = do.call("merge", append(list(def), def@externalClasses)) 

                file = file(file)
                open(file, "w")

		for(i in prepend) {
		   cat(c(paste("#", i), readLines(i), "\n"), sep = "\n", file = file)
      	        }
                on.exit(close(file))
                ns = callNextMethod()
              }

	       if(FALSE && length(def@events)) {
                   EventTableFile = "EventTable.rda"                 
                   env = createEventTable(def@events)
   	           save(list = objects(envir = env), envir = env,
                        file = EventTableFile)
               }


              if(!is.na(namespace) && length(namespace) > 0)  
   	   	  writeNamespace(ns, namespace)

              ns
           })


writeNamespace =
function(ns, nsFile = "NAMESPACE", copy = character())
{
 if(is.character(nsFile)) {
   nsFile = file(nsFile, open = "w")
   on.exit(close(nsFile))
 }

 cat("import(RDCOMClient)\n\n", file = nsFile)

 for(i in ns) {
  if(length(i$classes)) {
     cat("exportClasses(\n   ", paste(sQuote(unlist(i$classes)), collapse = ",\n   "), 
                "  )\n\n\n", file = nsFile)
  }

  if(length(i$variables))
     cat("export(\n   ", paste(sQuote(i$variables), collapse = ",\n   "), 
                "  )\n\n\n", file = nsFile)
 }

 cat("\n\nexportMethods('names', '$', '$<-', '[[', '[[<-', 'coerce')\n", file = nsFile) 

 if(length(copy)) {
    sapply(copy, copyFile, nsFile)
 }
}

##################################################################################
# Utilities
	
copyFile = 
function(filename, con)
{
 txt = readLines(filename)
 cat(txt, "\n", sep = "\n", file = con)
}


backquote = 
function(x)
  paste('`', x, '`', sep = "")			

sQuote = 
function(x)
  paste("'", x, "'", sep = "")			
##################################################################################

createEventTable = 
#
# This takes the definitions for the different interfaces
# and creates two objects which can then be saved.
# The two objects are 
#   1)  a named character vector mapping the human names for the interfaces
#       e.g. IWorkbookEvents to the corresponding UUID string.
#   2)  a named list indexed by the UUIDs in 1) and whose elements are the 
#       template functions that are used when creating an instance of an event handler/server.
#
#  The argument prefix is to give names to the variables.
# This could be the name of the library/application, e.g. Word or MSForms
# to separate them from other objects we programatically create.
	
function(eventDefs, prefix = "")
{
  uuidMap = sapply(eventDefs, function(x) x@guid)
  names(uuidMap) = sapply(eventDefs, function(x) x@interfaceName)

  EventTable = lapply(eventDefs, function(x) x@template)
  names(EventTable) = uuidMap

  return(list(UUIDMap = uuidMap, EventTable = EventTable))
  
  env = new.env()
  assign(paste(prefix, "UUIDMap", sep = ""), uuidMap, envir = env)
  assign(paste(prefix, "EventHandlerTable", sep = ""), EventTable, envir = env)

  env
}


setMethod("writeCode", c("RCOMTypeLibraryInterfaceDefinition"),
           function(def, file = stdout(), context = NULL,
	  	     prepend = character(),
                     enumsAsValues = TRUE, namespace = "NAMESPACE", ...) {

               ns = list()
  	       cat("library(RDCOMClient)\n\n", file = file)
               ns$Classes = writeClasses(def, file)
               ns$Enums = writeEnums(def, file, enumsAsValues)

               sapply(def@types,
                      function(x) {
                        writeCode(x, file = file, context, ...)
                      })

               if(length(def@coclass)) {
	          for(i in names(def@coclass)) {
                    cat("if(isClass('", def@coclass[i], "'))\n", file = file, sep = "")
                    cat("\tsetClass('", i, "', contains = c('CompiledCOMCoClass', '", def@coclass[i], "'), prototype = list(coclass = '",  def@coclass[i] , "'))\n", sep = "", file = file)
                    }
                }

               ns$Events = writeEventCode(def@events, file)

               ns
           })


# We need a way to represent information for defining a new class
# and it becomes a little trickier when we have to deal with
# COMTypedNamedList types since we need to specify the prototype.
# So, for the moment, we use two types to represent a class
# definition:  InlineClassDefinition and NamedClassDefinition
# and define a method 'classDefinitionString' which is used
# to get the output that is written in the generate code that
# defines the new class.
#
# InlineClassDefinition is for when we want to provide the 
# explicit and literal R command to define the new class.
# So the definition of this class is given by this code.

#
# NamedClassDefinition is the more symbolic version which 
# contains a collection of class names as a character vector
# and gives the name of the class being defined first and then
# the value of the 'contains' argument in the call to setClass.

setClass("ClassDefinition",
          representation(code = "character",
                         className = "character"))

setClass("InlineClassDefinition", contains = "ClassDefinition")
setClass("NamedClassDefinition", contains = "ClassDefinition")


setClass("ExternalClassDefinition",
	   representation(info = "ITypeInfo",
	                  library = "ITypeLib"),
	   contains = "NamedClassDefinition") # just convenient for the printing at present.
	

setGeneric("classDefinitionString",
             function(x, ...) {
                 standardGeneric("classDefinitionString")
             })


setMethod("classDefinitionString",
           "NamedClassDefinition",
             function(x, ...) {
              paste("setClass('", x@className, "', contains = c(", 
	                           paste("'", x@code[-1],"'", sep = "", collapse = ", "),
	                          "))", sep = "")
             }
        )


setMethod("classDefinitionString",
           "InlineClassDefinition",
             function(x, ...) {
              x@code
             }
        )


#
#
# This class represents information extracted and generated from
# the type library. It contains R code for each of the elements
# in the IDispatch type to access and set properties, and
# call methods.  It also contains a vector of two values giving
# the name of the new class in R and its parent class.
#
# expanded to have base class and extend for events and dispatch
#
setClass("RCOMTypeInterfaceDefinition",
           representation(className = "ClassDefinition", 
                          methods = "list",
                          guid = "character",
                          interfaceName = "character"
                          ))


setClass("RCOMEventInterfaceDefinition", 
            representation(constructor = "character",
                           template = "COMEventServerInfo",
                           constructorName = "character"),
            contains = "RCOMTypeInterfaceDefinition")

setClass("RCOMDispatchInterfaceDefinition",
               representation("RCOMTypeInterfaceDefinition",
                              propertyGets = "list",
                              propertySets = "list"))


#
# This is incomplete. The idea is that this will enable the 
# definitions directly in R.  It may do so by writing text
# and source()ing that back into R or explicitly parsing
# the different values, or generating the code directly.

setGeneric("define", function(x, where, ...) standardGeneric("define"))
setMethod("define", "RCOMTypeInterfaceDefinition",
            function(x, where, ...) {
              setClass(x@className[1]@className[1], contains = x@className[2])

              txt = "function(x, name) {}"
              setMethod("$", c(x@className@className[1], "character"),
                 eval(parse(text = txt), globalenv()))
            })



# Write the code for a particular ITypeInfo element to
# a file.
setMethod("writeCode", c("RCOMTypeInterfaceDefinition", "character"),
           function(def, file = stdout(), context = NULL, 
	  	     prepend = character(),
                    enumsAsValues = TRUE, namespace = "NAMESPACE", ...) {
              file = file(file)
              open(file, "w")
              on.exit(close(file))
	      writeCode(def, file, context, enumsAsValues, ...)
           })

functionListToString =
  #
  # Generates a string that gives the name = value pairs
  # for the entries in def where def is assumed to be
  # a list of character vectors, actually COMOperationDefinition
  # objects.
function(def)
{
   if(length(def) == 0)
    return("")

   def = def[order(names(def))]
	
   paste(paste("'", names(def), "'", sep = ""), 
         sapply(def, as, "character"), 
         sep = " = ", collapse = ",\n")
}

#
# We write the lists of functions as regular, top-level
# variables because we want to share them with the [[
# and the $ methods. Also, we want to be able to compute
# on them to provide help, etc.  This decision means
# we do not use closures to trap the value of x.
# As a result, we use the technique of inserting the value
# into the formal parameter list via a default value.
#
setMethod("writeCode", "RCOMTypeInterfaceDefinition",
           function(def, file = stdout(), context = NULL, 
	  	     prepend = character(),	
                     enumsAsValues = TRUE, namespace = "NAMESPACE", ...) {


            ids = createTypeVarName(def@className, 
                                    c("Methods", "SetProperty", "GetProperty", "$", "$<-"))


             cat(ids["GetProperty"], " = list(", file = file)
	     txt = functionListToString(def@propertyGets)
             cat(txt, ")\n", file = file)

             cat(ids["SetProperty"], " = list(", file = file)
	     txt = functionListToString(def@propertySets)
             cat(txt, ")\n", file = file)

             cat(ids["Methods"], " = list(", file = file)
	     txt = functionListToString(def@methods)
             cat(txt, ")\n", file = file)
           })
           

generateClass =
  #
  # Creates R code to access the methods and properties for
  # an ITypeInfo described in a type library.
  #
function(id, lib, defaultClass = "CompiledCOMIDispatch",
         defaultListClass = c("COMList", "COMTypedList"),
	 verbose = FALSE)
{
      if(verbose)
         cat("<generateClass>", id, "\n")

      info = lib[[id]]

      funcs = getFuncs(info, lib)

      className = computeClassName(info, id, lib, defaultClass, defaultListClass, funcs = funcs)

# Some PropertyGets functions take
# arguments.  Some take 1 or 2 arguments.
# for these functions, we want to add them to the
# collection of functions.

      propGets = funcs[sapply(funcs, is, "PropertyGetDescription")]

      w = sapply(propGets, function(x) length(x@parameters) == 0)
      if(length(w))
        gets = lapply(propGets[w], generatePropertyGet, info, funcs, lib, verbose = verbose) # funcs used to be propGets[w]
      else
        gets = list()

 # PropertySet.  (ignore the PropertySetRef for now)
      propSets = funcs[sapply(funcs, is, "PropertySetDescription")]

      sets = lapply(propSets, generatePropertySet, info, funcs, lib, verbose = verbose) # funcs used to be propSets

 # Regular methods + left overs from !w
 #XXX watch the -(1:7). Use real names in matching.
      methods = funcs[sapply(funcs, is, "FunctionInvokeDescription")]
      idx = match(BasicIDispatchMethodNames, names(methods))
      methods = methods[- idx[!is.na(idx)]]
      methods = lapply(methods, generateMethod,  info, lib, funcs)          

 # Generate the actual code from these elements.

    new("RCOMDispatchInterfaceDefinition", 
         className = className, propertyGets = gets, propertySets = sets, methods = methods)
}


convertDefaultValue = 
#XXX convert the default value to something meaningful.
function(var)
{
  NA
}



createSignature =

function(desc)
{
   # Determine which are optional arguments in the COM method.
  optional = sapply(desc@parameters,  function(var) var@style@optional)

   # Generate the signature/formal parameters for this R function.
  signature = 
     sapply(desc@parameters, 
             function(var) {
                if(!var@style@optional)
                  var@name
                else
                  paste(var@name, convertDefaultValue(var), sep = " = ")
             })

  list(signature = signature, optional = optional)
}  

# Can marry PropertyGet and Method since they inherit from FunctionDescription
# and have the same basic structure.
generateMethod =
function(desc, info, lib, funcs = getFuncs(info, lib))
{
  sig = createSignature(desc)
  
   # Define the function 
  txt = paste("function(",  paste(sig$signature, collapse = ", "),
               ifelse(length(sig$signature), ", ", ""), ".x){")

  if(any(!sig$optional)) {
   txt = c(txt,
          c(paste("if(", paste("missing(", names(desc@parameters)[!sig$optional], ")", collapse = "||"), ") {"),
            paste("stop('You must specify a value for the argument(s) ", 
                            paste(names(desc@parameters)[!sig$optional], collapse = ", "), "')"),
             "}"))
  }

   # define
  if(length(desc@parameters))
   txt = c(txt, 
           sapply(desc@parameters,
                  function(i) {
                   x = generateConvertArgumentCode(i@type, info, lib, i@name)
                   if(x != "" && sig$optional[i@name]) {
                     paste("if(!missing(", i@name, "))", x)
                   }
                   else
                     x
                  }))

  txt = c(txt, 
          paste(".args = rep(as.logical(NA), ", length(desc@parameters), ")"))

  if(length(desc@parameters))
   txt = c(txt,
           lapply(seq(along = desc@parameters),
                  function(i) {
                    c(paste("if(!missing(", desc@parameters[[i]]@name, "))"),
                      paste(".args[", i, "] = TRUE"))
                  }))

  comCall = paste("ans = .COM(.x, '", desc@name, "', ", 
                   paste(names(desc@parameters), collapse = ", "),
                   ifelse(length(desc@parameters), ", ", ""),
                   ".dispatch = as.integer(1), .ids =", desc@memid, 
                   ", .suppliedArgs = .args",
                   ")", sep = "", collapse = "")
                          
  txt = c(txt, comCall)

  txt = c(txt, generateConvertResultCode(desc@returnType, info, lib, "ans"), "}")

  if(length(desc@parameters))
     classes = lapply(desc@parameters, function(x) computeClassName(info, x@type, lib, funcs = funcs))
  else
     classes = list()

  classes = c(classes, computeClassName(info, desc@returnType, lib, funcs = funcs))

    # Turn this into a COMMethodDefinition so we know what it is and can use methods
    # to show it.
  new("COMMethodDefinition", unlist(txt),
        referencedClasses = classes)
}

generatePropertySet =
#
# Assuming there is only one parameters.
#
function(desc, el, funcs, lib, verbose = FALSE)
{
    txt = "function(x, value) {"
    txt = c(txt, generateConvertArgumentCode(desc@parameters[[1]]@type, el, lib, "value"))
    txt = c(txt, paste(".COM(x, '", desc@name, "', value, .dispatch = as.integer(4))", sep = ""))

    txt = c(txt, "}")

    refClasses = list(computeClassName(el, desc@parameters[[1]]@type, lib, funcs = funcs))
    if(is.null(refClasses[[1]])) refClasses = list()

    new("COMPropertySetDefinition", txt,
            referencedClasses = refClasses)

}

# vector of names of parameters which indicate that the
# value is to be treated as a file name and so had better be 
# normalized.
# Edit this to add more using 
#    assignInNamespace("FileNameParameterNames", newValue, "SWinTypeLibs")

FileNameParameterNames = c("FileName")

isFileParameter = 
# This is the function that checks the 
function(varName)
{
  varName %in% FileNameParameterNames
}

generateConvertArgumentCode =
function(type, info,  lib, varName)
{
  target = targetRType(type, info, lib)
  if(length(target) == 0 || target == "")
    return("")

  if(target == "character" && varName %in% FileNameParameterNames)
     varName = paste("utils::normalizePath(", varName, ")")

  ans = paste(" as(", varName, ", '", target,"')", sep = "")

  if(is(type, "TypeDescriptionRef")) { #XX May need to be more specific and more inclusive.
    if(names(type@reftype) %in% names(lib)) {
       rtype = lib[[names(type@reftype)]]
       if(is(rtype, "ITypeInfoEnum"))
         ans = paste("as.integer(", ans, ")")
    }
  }

  paste(varName, ans , sep = " = ")
}


# Type conversion method. This is used to determine the type to which
# an R value should be coerced, again to an R value, but one
# that is appropriate for a call to a COM method, etc.
# It can return "" in which case the value does not need to be coerced.
# It is okay as-is, e.g.  we don't know about the expected type, or
# it already is in the correct form (!)
setGeneric("targetRType", function(type, info, lib) standardGeneric("targetRType"))

setMethod("targetRType", "TypeDescription",
            function(type, info, lib) {
               w = match(type@name, names(PrimitiveTypeConverters))
	       if(!is.na(w))
                 return(PrimitiveTypeConverters[w])

#XXX should we check for COMIDispatch*
               ""
            })

#XXX See getRTypeName - it is identical
setMethod("targetRType", "TypeDescriptionRef",
            function(type, info, lib) {
                 val = getRefTypeName(info, type@reftype, lib)
                 if(is.na(val))
                   val = "COMIDispatch"

                 val
            })

generatePropertyGet =
#
# Create functions for the given PropertyGetDescription accessor
#
function(desc, el, funcs, lib, verbose = FALSE)
{
    txt = "function(x) {"
    txt = c(txt, paste("\t ans = .COM(x, '", desc@name, "', .dispatch = as.integer(2))", sep = ""))
    txt = c(txt, generateConvertResultCode(desc@returnType, el, lib, "ans"))
    txt = c(txt, "}")

    refClasses = list(computeClassName(el, desc@returnType, lib, funcs = funcs))
    if(is.null(refClasses[[1]]))  
      refClasses = list()

    new("COMPropertyGetDefinition", txt,
           referencedClasses = refClasses)
}

#
#
# Do we need to return auxillary information such as new
# methods, classes, function definitions to support the
# code generated here so it can be shared.
#

setGeneric("generateConvertResultCode",
            function(type, info, lib, varName = "ans") standardGeneric("generateConvertResultCode"))


# A basic mapping of the COM names to R types,  in the form COM = R.
# I had thought about using functions here for generality, or expressions,
# but we use just names for the present.
PrimitiveTypeConverters = c(
                   long = "integer",
                   float = "numeric",
                   "VARIANT_BOOL" = "logical",
                   "BSTR" = "character",
                   "double" = "numeric",
                   COMIDispatch = "COMIDispatch",
                   "<User Defined>" = "COMIDispatch"
#                   VARIANT = function(x) x,
                  )

setMethod("generateConvertResultCode",
           "TypeDescription",
	   function(type, info, lib, varName = "ans") {
	     if(type@name %in% names(PrimitiveTypeConverters))
               return(varName)

             paste("\t", varName)
})

setMethod("generateConvertResultCode",
           "TypeDescriptionRef",
	   function(type, info, lib, varName = "ans") {
	      name = getRefTypeName(info, type@reftype, lib)
                 # Need to check we are defining a type for the value of name
	      if(is.na(name)) # This means that the type is in a different type library.
                 name = "COMIDispatch"

            
              paste("\tif(!is.null(", varName, "))\n\t as(", varName, ", '", name, "')\n\telse\n\t   ", varName,sep = "")
})


#
# Used to determine what R class a COM type should extend.
# This merely determines whether the COM type is a COMList,
# a COMTypedList or just a generic COMIDispatch.
# The first two relate to whether there are Item and Count 
# methods. If the Item method returns an object whose
# type name is the singular form of the specified typeName
# argument, then we use a COMTypedList.

# When generating bindings, one would call this with a 
# value for defaultClass that provides an extension of COMIDispatch
# that is a virtual class for all types in the type library being
# processed.
setGeneric("computeClassName",
	    function(info, typeName, lib, 
	              defaultClass = "COMIDispatch",
		      defaultListClass = c("COMList", "COMTypedList"),
                      funcs = getFuncs(info, lib)) {
                   standardGeneric("computeClassName")
            })


setMethod("computeClassName",
            c("ITypeInfo", "TypeDescriptionRef"), 
	    function(info, typeName, lib, 
	              defaultClass = "COMIDispatch",
		      defaultListClass = c("COMList", "COMTypedList"),
                      funcs = getFuncs(info, lib)) {

              ref = typeName@reftype

              typeName = getRefTypeName(info, ref, lib)
	      if(is.na(typeName)) {
                return(NULL)
 	      }

                # The name may not be in this library as getRefTypeName()
	        # is looking into the library where the reference type is
	        # actually defined.
              if(!(typeName %in% names(lib))) {
                 info = getRefTypeInfo(info, ref, lib, follow = TRUE)
 	         lib = getTypeLib(info)
 	         funcs = getFuncs(info)

                 def = computeClassName(info, typeName, lib, defaultClass, defaultListClass, funcs)

	         return(new("ExternalClassDefinition", def, info = info, library = lib))
		
              } else
                 info = lib[[typeName]]

              #  callNextMethod()
              # This seems really slow.
              computeClassName(info, typeName, lib, defaultClass, defaultListClass, funcs)
            })

setMethod("computeClassName",
            c("ITypeInfo", "PointerTypeDescription"), 
	    function(info, typeName, lib, 
	              defaultClass = "COMIDispatch",
		      defaultListClass = c("COMList", "COMTypedList"),
                      funcs = getFuncs(info, lib)) {

                NULL
            })

setMethod("computeClassName",
            c("ITypeInfo", "TypeDescription"), 
	    function(info, typeName, lib, 
	              defaultClass = "COMIDispatch",
		      defaultListClass = c("COMList", "COMTypedList"),
                      funcs = getFuncs(info, lib)) {

               #XXX Check these only refer to the primitive types in this context.
                NULL
            })



setMethod("computeClassName", 
            c("ITypeInfo", "character"),
function(info, typeName, lib, 
          defaultClass = "COMIDispatch",
          defaultListClass = c("COMList", "COMTypedList"),
          funcs = getFuncs(info, lib))
{
  if(is.na(typeName)) {
      uuids = getTypeLibUUIDs(lib)
      idx = match(info@guid, uuids)
      typeName = names(uuids)[idx]
  }

     # See if this name really corresponds to an IDispatch 
     # element within a CoClass object within the library.
  if(typeName %in% names(lib)) {
      el = lib[[typeName]]
      if(is(el, "ITypeInfoCoClass")) { 
         els = getElements(el)
         if(paste("_", typeName, sep = "") %in% names(els))
          typeName = paste("_", typeName, sep = "")
      }
  }

  # Check if this is a COMList or COMTypedList.
  # Otherwise, return defaultClass


  if(is.function(defaultListClass))
     className = defaultListClass(funcs, typeName, info, lib)
  else
     className = getListClassName(funcs, typeName, info, lib, defaultListClass, defaultClass)

  if(!is.null(className)) {
    className
  } else
    new("NamedClassDefinition", code = c(typeName, defaultClass), className = typeName)
}
)

NameElementNames = c("Name", "NameLocal")

getListClassName =
#
# A basic function that checks the collection of functions
# to see if this is an OfficeList and determines what the
# class definition should be.
#XXX  needs to provide more information about COMTypedNamedList.
#
function(funcs, typeName, info, lib,
           defaultListClass = c("COMList", "COMTypedList"),
            defaultClassName = "CompiledCOMIDispatch")
{
  if(length(typeName) == 0)
   stop("getListClassName called with empty typeName")

  if(!any(is.na(match(c("Item", "Count"), names(funcs))))) {
	 # what does Item return
     type = funcs$Item@returnType
     if(is(type, "TypeDescriptionRef")) {
               # this is the type returned by Item
        name = getRefTypeName(info, type@reftype, lib)
	if(is.na(name))
	  return(NULL)

	        # Is it a singular version of the current container class,
	        # e.g. Document and Documents
           #XXX Is this an necessary restriction. Added TRUE to see!
        if(TRUE || paste(name, "s", sep = "") == typeName)  {

            # If this (lib[[name]]) is a CoClass object, then 
            # we have to use a different name.

          rtype = lib[[name]]
          if(is(rtype, "ITypeInfoCoClass"))
              # resolve the referenced class. XXX currently just first element.
             rtype = getElements(rtype)[[1]]

	
          if(is(rtype, "ITypeInfoDispatch")) {

	       # See if we should declare this as a COMTypedNamedList
	       #  Find the names of the methods & attributes in this singular
	       #  form.
             elNames = names(getElements(rtype))
# old
#      w = grep(paste(name, "$", sep = ""), elNames, perl = TRUE, value = TRUE)
#     if(length(w))....
             nameId = structure(NameElementNames %in% elNames, names = NameElementNames)
             if(any(nameId)) {
               isParameterized = !nameId["Name"]
               listClassName = if(!isParameterized) "COMTypedNamedList" else "COMTypedParameterizedNamedList"
               
               txt = paste("setClass('", 
                           typeName, 
                           "', contains = c(", paste("'", c(listClassName, defaultClassName), "'", sep = "", collapse = ", "), ")",
       	                   ", prototype = list(name = '",  name,  "'",  if(isParameterized) paste(", nameProperty = '", (NameElementNames[nameId])[1], "'", sep = ""), ")",
 	                  ")",
                           sep = "")

	       return(new("InlineClassDefinition", code = txt, className = typeName))                 
             }
          }

          return(new("NamedClassDefinition", code = c(typeName, c(rev(defaultListClass)[1], defaultClassName)), className = typeName))
	}
      }

     return(new("NamedClassDefinition", code = c(typeName, defaultListClass[1]), className = typeName))
   }
}


BasicIDispatchMethodNames =
  c("QueryInterface", "AddRef", "Release", "GetTypeInfoCount",
    "GetTypeInfo", "GetIDsOfNames", "Invoke")            



generateEventInterface =
  # We want a class name for the new class
  # and a constructor function that allows the user
  # to create a server
function(info, id, lib, defaultClassName = "CompiledCOMServer",
          verbose = FALSE)
{
  if(verbose)
      cat("<generateEventInterface>", id, "\n")

  funcs = getFuncs(info, lib)
    # Discard the standard, built-in IDispatch methods
  w = match(BasicIDispatchMethodNames, names(funcs))
  if(length(w) && !all(is.na(w)))
    funcs = funcs[ - which(!is.na(w)) ]

  
  defs = RDCOMEvents::createFunctionTemplates(funcs, RDCOMEvents::createCompiledFunction, lib) # lib is part of ...

  funcName = paste(id, "Handler", sep = "")
  constructor = paste(funcName,
                      " = function(methods) {\n",
                      "  uuid = '", info@guid, "'\n",
                      "  template = EventHandlerTable[[uuid]]\n",
                      "  createEventServerFromTemplate(methods, template, '", id, "')\n",
                      "}\n", sep = "")


#XXX - fix this . repetitive CompiledCOMServer.
# Does this still happen? Shouldn't as we have shifted to a new mechanism.
  className = computeClassName(info, id, lib, defaultClassName, funcs)
#  if(is.character(className)) {
#      className = c(className, defaultClassName)[1:2]
#  }

  template = RDCOMEvents::createCOMEventServerInfo(info, .lib = lib, complete = TRUE)

  new("RCOMEventInterfaceDefinition", className = className, 
                                      methods = defs, 
                                      guid = info@guid, 
                                      interfaceName = id,
				      template = template,
                                      constructor = constructor, 
                                      constructorName = funcName
                                    )

}


# Is this the same as targetRType. Not quite, but could be made to be.
setGeneric("getRTypeName", function(type, info, lib) standardGeneric("getRTypeName"))

setMethod("getRTypeName", "TypeDescription",
           function(type, info, lib) {
              targetRType(type, info, lib)
           })

setMethod("getRTypeName", "TypeDescriptionRef",
          function(type, info, lib) {
                # Should we call computeClassName() here?
              if(type@name == "void")
                return(character())

              val = getRefTypeName(info, type@reftype, lib)
              if(is.na(val))
                val = "COMIDispatch"
             
              val
          })


generateEventHandler =
  #
  # This creates a simple function with no or a degenerate body
  # but that has the appropriate signature so that it can be
  # called with no effect. This is needed when implementing
  # an event handler.  However, we should create an event
  # handler infrastructure (if we haven't already - see RDCOMEvents) that can
  # handle methods with no associated R implementation.
  # e.g. if method1 is not provided, it just handles the IDispatch::Invoke()
  # method for that as a no-op.  
  # This automatically happens at present (with a warning).

  #
  # All this function needs to do is generate the appropriate
  # signature and return. If the return type is not void,
  # it should return a value of the appropriate type.
#XXX
  #  For example, for the event interface ApplicationAddRef, Release, etc. return ULONG.
  #  Of course, we don't implement these methods so these disappear.
  # Are there other methods that have a non-void return type?
function(desc, info, lib)
{
   # Put type information about the different arguments into the methods.  
  sig = createSignature(desc)

  types = sapply(desc@parameters, function(x) targetRType(x@type, info, lib)) #getRTypeName(x@type, info, lib))
  
  txt = paste("function(",  paste(sig$signature, collapse = ", "), "){")

  txt = c(txt, paste("# ", names(types), types)) 

  
  if(desc@returnType@name != "void") {
    warning("Non-void return type", desc@returnType@name,"\n")
  }

  txt = c(txt, "}")

  new("COMEventDefinition", txt)
}  






mergeInterfaces = 
 function(x, y, ...) {
              
             for(y in x@externalClasses) {
               x@types = append(x@types, y@types)
               x@enums = append(x@enums, y@enums)
               x@events = append(x@events, y@events)
               x@externalClasses = append(x@externalClasses, y@externalClasses)
             }         

             x
           }

if(FALSE) {
if(!isGeneric("merge"))
setGeneric("merge", function(x, y, ...) standardGeneric("merge"))
setMethod("merge",
           c("RCOMTypeLibraryInterfaceDefinition"),
           function(x, y, ...) {
 
             for(y in c(y, list(...))) {
               x@types = append(x@types, y@types)
               x@enums = append(x@enums, y@enums)
               x@events = append(x@events, y@events)
               x@externalClasses = append(x@externalClasses, y@externalClasses)
             }         

             x
           })
}


setMethod("dput", "ANY",
function(x, file = "", control = c("keepNA", "keepInteger", "showAttributes"))
{
          if(isS4(x))
            dput.S4(x, file, control)
          else
            base::dput(x, file, control)
})



setMethod("dput", "raw",
            function(x, file = "", control = c("keepNA", "keepInteger", "showAttributes")) {
                  cat("as.raw(", file = file)
                  dput(as.numeric(x), file)
                  cat(")\n", file = file)
            })

# An S4 object.
dput.S4 =
function(obj, file = stdout(), control = c("keepNA", "keepInteger", "showAttributes"))
{
  cat("new(", sQuote(class(obj)), ",\n\t\n", file = file)
  slots = names(getSlots(class(obj)))
  sapply(seq(along = slots),
          function(i) {
           cat("\t", sQuote(slots[i]), "= ", file = file)
               # Could write to a text connection and then gsub new lines with \n\t
           con = textConnection("bob", "w", local = TRUE)
           on.exit(close(con))
           dput(slot(obj, slots[i]), con, control)
           close(con); on.exit()
           cat(gsub("\\\n", "\\\n\\\t", bob), file = file)
           if(i < length(slots))
              cat(",\n\t", file = file)
         })
  cat("\t)\n", file = file)
}


writeEventCode =
  #
  #  We need the name of the class for which each set of events applies
  #  Then we write a method for addEventHandler
  
function(events, file = stdout(), useNamespace = TRUE)
{
  cat("library(RDCOMEvents)\n", file = file)
  ns = lapply(names(events),
              function(id) {
               writeClassEvent(id, events[[id]], file)
              })
  ns
}  

writeClassEvent =
function(className, eventInfo, file = stdout())
{
  cat("\n# addDCOMEventHandler method for", className, "\n", file = file)
  cat("setMethod('addDCOMEventHandler', c(", sQuote(className), "),\n", file = file)
  cat("function(obj, ..., \n",
      "\t.methods = list(...),\n",
      "\t.connectionPoint = NA,\n",
      "\t.uuidMap = NA,\n",
      "\t.eventTable = list())\n{\n" , file = file)

  cat("  if(is.na(.uuidMap))\n.uuidMap =", file = file)
  map = structure(sapply(eventInfo, function(x) x@guid), names =names(eventInfo))
  con = textConnection("bob", "w", local = TRUE)
  dput(map, file = con)
  close(con)
  cat(gsub("\\\n", "\\\n\\\t", bob), "\n\n", file = file)

  cat("  if(length(.eventTable) == 0)\n\t.eventTable = structure(list(", file = file)
  #dput(eventInfo, file = file)
  n = length(eventInfo)
  sapply(seq(along = eventInfo), function(i) {
                        cat("\t\t", file = file)
                           # We need to geerate something that is of a class we can subsequently instantiate at run time.
                           # We can move the class RCOMEventInterfaceDefinition into RDCOMEvents and serialize as is.
                           # Alternatively, we could serialize in a customized manner to create the event method description
                           # we want. We don't need the constructor, constructorName or the methods, just the
                           #    template, guids, perhaps the interface name
                        # dput.S4(eventInfo[[i]], file)
                        writeEventDefinition(eventInfo[[i]], file)
                        if(i < n)
                           cat(",\n", file = file)
                        })
  cat("\t\t), .Names = c(", paste(dQuote(names(eventInfo)), collapse = ", "), "))\n", file = file)
#  dput.S4(eventInfo, file = file)    
  
  cat("\n\naddDCOMEventHandler(obj, .methods = .methods, .connectionPoint = .connectionPoint, .uuidMap = .uuidMap, .eventTable = .eventTable)", file = file)
  cat("\n})\n\n", file = file)
#  dput()
}

setGeneric("writeEventDefinition",
           function(obj, file = stdout(), indent = "\t\t")
                standardGeneric("writeEventDefinition"))


setMethod("writeEventDefinition",
           "RCOMEventInterfaceDefinition",
            function(obj, file = stdout(), indent = "\t\t") {

              #XXX will be removed. We change RDCOMEvents to expect the quote and to generate the quote.
              obj@template@methods = lapply(obj@template@methods, fixParameterContracts)
              
               cat("new('CompiledCOMEventInterfaceInfo', \n", file = file)
               cat(indent, "template = ", file = file)
               dput(obj@template, file = file)
               cat(",\n", indent, "guid = '", obj@guid, "',\n", sep = "", file = file)
               cat(indent, "interfaceName = '", obj@interfaceName, "')", file = file)
            })

fixParameterContracts =
function(f)
{
  tracts = attr(f, "ParameterContracts")
  tracts = lapply(tracts,
                   function(e) {
                     o = call("quote")
                     o[[2]] = e
                     o
                   })
  attr(f, "ParameterContracts") = tracts
  f
}  
