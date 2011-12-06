if(FALSE) {
.Call =
function(name, ...)
{
  base::.Call(name, ...)
}
}	
	
library(methods)
setClass("ITypeLib", representation(ref="externalptr",
                                    fileName = "character")
#                       prototype = list(fileName = NA)
         )

setClass("IContainingTypeLib", contains = "ITypeLib")
setClass("INamedTypeLib",
           representation(elNames = "character",
                          infoEls = "list",
                          uuids = "character"
                         ),
           contains = "ITypeLib")

# Information extracted from a type library via getTypeLibInfo().
setClass("ITypeLibIdentifier",
         representation(guid = "character",
                        version = "integer",
                        lcid = "numeric",
                        flags = "integer",
                        syskind = "integer"))




setClass("ITypeInfo", 
	  representation(ref="externalptr", type="integer", guid="character"),
          prototype=list(type=as.integer(-1), guid=""))


setClass("ITypeInfoEnum", representation("ITypeInfo"))
setClass("ITypeInfoRecord", representation("ITypeInfo"))
setClass("ITypeInfoModule", representation("ITypeInfo"))
setClass("ITypeInfoInterface", representation("ITypeInfo"))
setClass("ITypeInfoDispatch", representation("ITypeInfo"))
setClass("ITypeInfoCoClass", representation("ITypeInfo"))
setClass("ITypeInfoAlias", representation("ITypeInfo"))
setClass("ITypeInfoUnion", representation("ITypeInfo"))


	

# No longer needed
# ITypeLibEntry is in fact ITypeInfo
#  setClass("ITypeLibEntry", representation("ITypeInfo"))
#	

setClass("TypeDescription",
          representation(name = "character")
        )

setClass("PointerTypeDescription",
          representation("TypeDescription",
	                 depth = "integer"),
	    prototype=list(name="", depth=integer(1)))

setClass("TypeDescriptionRef",
          representation("TypeDescription",
	                 reftype="numeric"),
	    prototype=list(name="", reftype=numeric(1)))	

setClass("ParameterStyle",
          representation( In = "logical",
                          out = "logical",
                          lcid = "logical",
                          retval = "logical",
                          optional = "logical"),
   prototype(In=FALSE, out=FALSE, lcid=FALSE, retval=FALSE, optional=FALSE))

setClass("ParameterDescription",
	    representation(name = "character",
	                   type = "TypeDescription",
                           style = "ParameterStyle",
			  # optional = "logical",
			   defaultValue="ANY"),
                           prototype = list(name = "", type = NULL,
                                            defaultValue = NULL))



setClass("ElementDescription",
          representation("VIRTUAL", 
                         name="character", 
                         hidden="logical",
	                 memid = "integer"))				   

setClass("VariableDescription",
          representation("ElementDescription",
                         type="TypeDescription"))
	
setClass("FunctionDescription",
          representation("ElementDescription",
	                 returnType= "TypeDescription",
                         parameters = "list",
			 invokeType = "integer",
	                 kind = "integer"),
          prototype = list(name="", returnType = NULL, parameters = NULL, optionalParameters = NULL, invokeType = integer(1), kind = integer(0)))


setClass("PropertyGetDescription", representation("FunctionDescription"))
setClass("PropertySetDescription", representation("FunctionDescription"))
setClass("PropertySetRefDescription", representation("PropertySetDescription"))
setClass("FunctionInvokeDescription", representation("FunctionDescription"))
	
setGeneric("getEnum", function(x) standardGeneric("getEnum"))
setMethod("getEnum", "ITypeInfoEnum",
           function(x) {
              els <- .Call("R_getEnums", x, PACKAGE="SWinTypeLibs")
	      if(!is.null(els))
                els <- unlist(els)
              els
           })


# This is very, very similar to getEnum. It is more general
# but seems to work in Enumeration types in Excel only.
setGeneric("getVars", function(x) standardGeneric("getVars"))
setMethod("getVars", "ITypeInfo",
           function(x) {
              els <- .Call("R_getCVars", x, PACKAGE="SWinTypeLibs")
	      if(!is.null(els))
                names(els) <- sapply(els, function(x) x@name)
              els
           })


	
setGeneric("getFuncs", function(x, lib = NULL) standardGeneric("getFuncs"))
setMethod("getFuncs", "ITypeInfo",
           function(x, lib = NULL) {
              els <- .Call("R_getCFuncs", x)
	      if(!is.null(els)) {
                if(is.null(lib))
                    lib = getTypeLib(x)
                invokes = .Call("R_getInvokeEnum")	        
                els <- lapply(els, processITypeInfoFunc, lib, x, invokes)
                names(els) <- sapply(els, slot, "name")
              }
              els
           })

processITypeInfoFunc =
function(el, lib, x, invokes = .Call("R_getInvokeEnum")) {
      # identify which type of element invocation this is,
      # e.g. func or a property get or set
     names(el@invokeType) <- names(invokes)[el@invokeType == invokes]
     nms = names(el@parameters)
     el@parameters = 
        lapply(seq(along = el@parameters),
                function(id) {
                  z = el@parameters[[id]]
                  z@name = nms[id]
                  if(is(z@type, "TypeDescriptionRef")) 
                     names(z@type@reftype) = getRefTypeName(x, z@type@reftype, lib)

        	  z
                })
     names(el@parameters) = nms
     el
 }


getRefTypeName =
function(info, ref, lib = NULL, follow = !is.null(lib))
{
  if(is(ref, "TypeDescriptionRef")) 
    ref = ref@reftype

  rinfo = getRefTypeInfo(info, ref, lib, follow = follow)

  if(is.null(rinfo))
      return(as.character(NA))
	
  return(getTypeDocumentation(rinfo, 0)["name"])

# Old way
#  uuids = getTypeLibUUIDs(lib)
#  names(uuids)[match(rinfo@guid, uuids)]
}


mapUUIDToName =
function(uuids, lib)
{
  ids = getTypeLibUUIDs(lib)
  names(ids)[match(uuids, ids)]
}


getTypeByUUID =
function(uuid, lib)
{
  uuids = getTypeLibUUIDs(lib)
  lib[[match(uuid, uuids)]]
}


setMethod("getFuncs", "COMIDispatch",
           function(x, lib = NULL) {
                getFuncs(getTypeInfo(x)[[1]], lib = lib)
           })


setMethod("[[", c("ITypeLib", "numeric"),
	   function(x, i, j, ...) {
	    getTypeInfo(x, i)[[1]]
  	})

setMethod("[", c("ITypeLib","numeric"),
	   function(x, i, j, ..., drop = TRUE) {
  	       getTypeInfo(x, i)
  	})


setMethod("[", c("ITypeLib","logical"),
	   function(x, i, j, ..., drop = TRUE) {
            # i = (1:length(lib))[i]
	    # x[i, drop = drop]

            x[which(i), drop = drop]
  	})		




setMethod("[[", c("ITypeLib", "character"),
            function(x, i, j, ...) {
	       tryCatch(.getTypeLibElements(x, i, drop = TRUE),
                      # If the expression fails, try treating the 
                      # id as a UUID and look it up that way.
                         error = function(err) getTypeByUUID(i, x))
            })

setMethod("[", c("ITypeLib", "character"),
            function(x, i, j, ..., drop = TRUE) {
	       .getTypeLibElements(x, i)
            })

setMethod("$", c("ITypeLib"), # "character"),
            function(x, name) {
	       .getTypeLibElements(x, name, drop = TRUE)
            })


if(FALSE) {
if(!isGeneric("lapply")) 
  setGeneric("lapply", function(X, FUN, ...) standardGeneric("lapply"))

if(!isGeneric("sapply")) 
 setGeneric("sapply", function(X, FUN, ..., simplify = TRUE, USE.NAMES = TRUE) standardGeneric("sapply"))
}
	
setMethod("lapply", "ITypeLib",
            function(X, FUN, ...) {
              v = lapply(names(X),
                          function(id)
                             FUN(X[[id]], ...))

	      names(v) = names(X)
	      v
            })


setMethod("sapply", "ITypeLib",
  function (X, FUN, ..., simplify = TRUE, USE.NAMES = TRUE) 
  {
    FUN <- match.fun(FUN)
    v = base::lapply(names(X),
                function(id)
                         FUN(X[[id]], ...))
    names(v) = names(X)
    answer = v
#XXX temporarily moved    answer <- lapply(X, FUN, ...)
    if (USE.NAMES && is.character(X) && is.null(names(answer))) 
        names(answer) <- X
    if (simplify && length(answer) && length(common.len <- unique(unlist(lapply(answer, 
        length)))) == 1) {
        if (common.len == 1) 
            unlist(answer, recursive = FALSE)
        else if (common.len > 1) 
            array(unlist(answer, recursive = FALSE), dim = c(common.len, 
                length(X)), dimnames = if (!(is.null(n1 <- names(answer[[1]])) & 
                is.null(n2 <- names(answer)))) 
                list(n1, n2))
        else answer
    }
    else answer
 })



setMethod("show", "INamedTypeLib",
            function(object) {
                show(as(object, "ITypeLib"))
            })

if(!isGeneric("names"))
setGeneric("names", function(x) standardGeneric("names"))

setMethod("names", "INamedTypeLib", function(x) x@elNames)

setGeneric("getTypeLibUUIDs", function(lib) standardGeneric("getTypeLibUUIDs"))

setMethod("getTypeLibUUIDs", "INamedTypeLib", function(lib) lib@uuids)

setMethod("[[", c("INamedTypeLib", "numeric"),
	   function(x, i, j, ...) {
	     x@infoEls[[i]]
  	})


setMethod("[", c("INamedTypeLib", "numeric"),
	   function(x, i, j, ..., drop = TRUE) {
	     x@infoEls[i]
  	})

setMethod("[[", c("INamedTypeLib", "character"),
	   function(x, i, j, ...) {
             idx = match(i, x@elNames)
             if(is.na(idx))
                stop("No such element(s) ", i, " in type library")
             x@infoEls[[idx[1]]]
  	})

#XXX Problems with this and other methods. The method table (via showMethods)
# reports that it is inheriting the methods from ITypeLib rather than the ones
# we have defined here!  And this is being borne out for this particular method.
# If we use the .Call() function in the package and trace access to 
# R_getTypeLibInfoEntry, we see that it is calling the method for ITypeLib
# even though we have an INamedTypeLib and a character vector.
setMethod("[", c("INamedTypeLib", "character"),
	   function(x, i, j, ..., drop = TRUE) {
             idx = match(i, x@elNames)
             if(any(is.na(idx)))
                stop("No such element(s) ",  paste(i[is.na(idx)], collapse = ", "), " in type library")
             x@infoEls[idx]
  	})





setMethod("[[", "ITypeInfoCoClass",
            function(x, i, j, ...) {
               getElements(x)[[i]]
            })

setMethod("[[", c("ITypeInfoCoClass", "numeric"),
	   function(x, i, j, ...) {
             getElements(x)[[i]]
     	   })


setMethod("[", "ITypeInfoCoClass",
            function(x, i, j, ..., drop = TRUE) {
               x = getElements(x)
	       callNextMethod()
            })






