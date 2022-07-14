setGeneric("LoadTypeLib", 
             function(fileName) standardGeneric("LoadTypeLib"))

setMethod("LoadTypeLib", "character",
function(fileName)
{
  if(!file.exists(fileName))
    stop("No such file")	 

  fileName = as.character(path.expand(fileName))
  v = .Call("R_loadTypeLib", fileName, PACKAGE="SWinTypeLibs")

  v@fileName = fileName

  v
})

# Non-exported function
loadLib =
function(fileName) {
    lib = getTypeLib(fileName)
    attr = getTypeLibInfo(lib) 
    LoadRegTypeLib(attr)
}

setMethod("LoadTypeLib", "ITypeLibIdentifier",
function(fileName)
{
    LoadRegTypeLib(fileName@guid, fileName@version, fileName@lcid)
})


setMethod("LoadTypeLib", "IContainingTypeLib", 
            function(fileName) {
               details = getTypeLibInfo(fileName)
               LoadTypeLib(details)
            })
	
setMethod("LoadTypeLib", "ITypeInfo", loadLib)

setMethod("LoadTypeLib", "COMIDispatch", loadLib)


setGeneric("queryTypeLibPath", function(guid, version, lcid = 0) standardGeneric("queryTypeLibPath"))



# Non-exported function
queryLib =
function(guid, version, lcid = 0) {
    lib = getTypeLib(guid)
    attr = getTypeLibInfo(lib) 
    queryTypeLibPath(attr)
}

setMethod("queryTypeLibPath", "COMIDispatch", queryLib)
setMethod("queryTypeLibPath", "ITypeInfo", queryLib)

setMethod("queryTypeLibPath", "ITypeLib",  
             function(guid, version, lcid = 0) {
                 queryTypeLibPath(getTypeLibInfo(guid))
             })


setMethod("queryTypeLibPath", "ITypeLibIdentifier", 
            function(guid, version, lcid = 0)
              queryTypeLibPath(guid@guid, guid@version, guid@lcid))

setMethod("queryTypeLibPath", c("character", "integer"), 
	   function(guid, version, lcid = 0) {
              .Call("R_QueryPathOfRegTypeLib", guid, version, 
                          as.numeric(lcid), PACKAGE = "SWinTypeLibs")	
    	   })



LoadRegTypeLib <-
function(guid, version, lcid = 0)
{
  if(is(guid, "ITypeLibIdentifier")) 
    return(LoadRegTypeLib(guid@guid, guid@version, guid@lcid))

  if(length(version)  < 2)
     stop("version must have both major and minor numbers")

  a = .Call("R_loadRegTypeLib", as.character(guid), as.integer(version), as.numeric(lcid), 
              PACKAGE = "SWinTypeLibs")

  a@fileName = queryTypeLibPath(a)

  a
}


setMethod("length", "ITypeLib", function(x) {
          .Call("R_getTypeLibInfoCount", x, PACKAGE="SWinTypeLibs")
})


setGeneric("getTypeInfo", function(obj, which = 1) standardGeneric("getTypeInfo"))

setMethod("getTypeInfo", c("ITypeInfo", "TypeDescriptionRef"),
             function(obj, which = integer(0))  {
                .Call("R_getTypeInfoHRefType", obj, which@reftype, PACKAGE="SWinTypeLibs")
             })

setMethod("getTypeInfo", c("ITypeLib", "numeric"),
             function(obj, which = integer(0))  {
              if(length(which) == 0) 
                which = 1:length(obj)

  	      vals <- .Call("R_getTypeLibInfoEntry", obj, as.integer(which-1), PACKAGE="SWinTypeLibs")

	      names(vals) <- names(obj)[which]
	      vals
	  })


	
setMethod("getTypeInfo", c("ITypeLib", "missing"),
             function(obj, which = integer(0)) 
  	      getTypeInfo(obj, 1:length(obj))
	  )

setMethod("getTypeInfo", c("ITypeLib", "character"), 
             function(obj, which = 1) { 
	         .getTypeLibElements(obj, which, drop = TRUE)
    	     }
	  )	

setMethod("getTypeInfo", "COMIDispatch", 
             function(obj, which = integer(0))  {
	       if(length(which) == 0)	   
	         which = 1:(.Call("R_getDCOMInfoCount", obj@ref, PACKAGE="SWinTypeLibs"))

  	      ans = .Call("R_getDCOMInfoEntry", obj@ref, as.integer(which-1), PACKAGE="SWinTypeLibs")
	      class(ans) = c("ListOfITypeInfo", "list")
	      ans
             })


	
TypeKinds <- 
 c(enum=0, record=1, module=2, interface=3, dispatch = 4,
  coclass = 5, alias = 6, union = 7)
storage.mode(TypeKinds) <- "integer"

getTypeLibTypes <-
#
# This seems to fail, but do we need it any more?
# It is just an efficient way of doing things
# rather than 
#     sapply(getTypeInfo(lib), function(x)x@type)
#
function(lib, which = 1:length(lib), byName = TRUE) 
{
  if(is.character(which)) {
   orig = which
   which <- match(which, names(lib))
   if(any(is.na(which)))
     stop("No such entries in type library" , paste(orig[is.na(which)], collapse=", "))
  }

  els <- .Call("R_getTypeInfoTypes", lib, as.integer(which - 1), PACKAGE="SWinTypeLibs")
  names(els) <- names(TypeKinds)[els + 1]

  if(byName) {
    els <- names(els)
    names(els) <- names(lib)[which]
  }

  els
}

#Generic defined in generate.S
setMethod("getTypeLibUUIDs", "ITypeLib",
function(lib)
{
  ans = .Call("R_getTypeLibUUIDs", lib, PACKAGE = "SWinTypeLibs")
  names(ans) = names(lib)
  ans
})


.getTypeLibElements <-
function(x, i, drop = TRUE, indices = FALSE)
{		
   idx = match(i, names(x))
   if(any(is.na(idx))) {
     gids = getTypeLibUUIDs(x)
     idx[is.na(idx)] = match(i, gids)
   }

   if(any(is.na(idx)))
     stop("No such element in the type library", paste(i[is.na(x)], collapse=", "))

   if(indices)
      return(idx)
	
   els <- vector("list", length(idx))
   els <- getTypeInfo(x, idx)
   names(els) = names(x)[idx]

   if(length(els) == 1 && drop) {
    els <- els[[1]]
   }
   els
}
	

setGeneric("getTypeDocumentation", function(lib, which) standardGeneric("getTypeDocumentation"))
setMethod("getTypeDocumentation", "ITypeLib",
	    function(lib, which) {
             .Call("R_getTypeLibDocumentation", lib, as.integer(which - 1), TRUE, PACKAGE="SWinTypeLibs")
    	    })

setMethod("getTypeDocumentation", "ITypeInfo",
	    function(lib, which) {
             .Call("R_getTypeLibDocumentation", lib, as.integer(which - 1), FALSE, PACKAGE="SWinTypeLibs")
    	    })


setMethod("names", "ITypeLib", function(x) .Call("R_getTypeLibNames", x, PACKAGE="SWinTypeLibs"))


setGeneric("getElements",
             function(x, recursive = FALSE) {
                 standardGeneric("getElements")
             })

#XXX
setMethod("getElements", "INamedTypeLib", function(x, recursive = FALSE) x@infoEls)
	
setMethod("getElements", "ITypeLib",
             function(x, recursive = FALSE) {
                x[1:length(x)]
             })


setMethod("getElements", "ITypeInfoAlias",
         function(x, recursive = FALSE) {
           o = .Call("R_getAlias", x, PACKAGE = "SWinTypeLibs")
	   if(is(o, "TypeDescription"))
	     return(o)

  	   if(is(o, "TypeDescriptionRef"))
              .Call("R_getRefTypeInfo", x, o@reftype, PACKAGE = "SWinTypeLibs")
	   else {
	    warning("potential problem here")
   	    o
   	  }
	#XXX need to deal with recursive argument.
  	 }
	)

setMethod("getElements", "ITypeInfoCoClass",
             function(x, recursive = FALSE) {
	         typeInfo = x
	         ans <- .Call("R_getCoClass", typeInfo, PACKAGE="SWinTypeLibs")
  	         if(recursive) 
		   ans <- lapply(ans, getElements)

		   	# we get the (names of the) library from the typeInfo
			# object to avoid having to ask the user to supply the lib.
	         uuids = getTypeLibUUIDs(getTypeLib(typeInfo))
	         guids = sapply(ans, function(x) x@guid)
	         names(ans) = names(uuids)[match(guids, uuids)]

		 ans
          })


setMethod("getElements", "ITypeInfo",        
  	  function(x, recursive = FALSE) {
	         typeInfo = x
		 type = names(typeInfo@type)
	 
		 if(type == "enum")
		     getEnum(typeInfo)
  	         else if(type == "dispatch" || type == "interface")
		     getFuncs(typeInfo)
 	         else
		    stop(type, " currently unsupported")
})


setGeneric("isSame", function(x, y) standardGeneric("isSame"))

setMethod("isSame", c("ITypeLib", "ITypeLib"),
                     function(x, y) {
                        x@fileName == y@fileName
                     })

setGeneric("getRefTypeInfo", function(info, href, lib = NULL, follow = !is.null(lib)) 
               standardGeneric("getRefTypeInfo"))

setMethod("getRefTypeInfo", c("ITypeInfo", "numeric"),
               function(info, href, lib = NULL, follow = !is.null(lib)) { 
                    # Get the ITypeInfo associated with this reference.
	          i = .Call("R_getRefTypeInfo", info, href)

		  if(i@type < 0)
	            return(NULL)

                    # If we have a library and want to follow to see if this
                    # reference is in a different library.
                  if(follow && is(lib, "ITypeLib")) {
                    olib = getTypeLib(i)
                    if(!isSame(lib, olib)) {
                         # So the ITypeInfo is in a different library.
                         # Resolve it to the one from that library.
                         olib = LoadTypeLib(olib)
                         id = getTypeDocumentation(i, 0)["name"]
                         i = olib[[ id ]]
                    }
                  }
                  i
	       })
setMethod("getRefTypeInfo", c("ITypeInfo", "TypeDescriptionRef"),
               function(info, href, lib = NULL, follow = !is.null(lib)) {
	         getRefTypeInfo(info, href@reftype, lib, follow)
	       })	



setGeneric("getTypeLib", function(x) standardGeneric("getTypeLib"))

setMethod("getTypeLib", "character",
              function(x)
                  LoadTypeLib(x))


setMethod("getTypeLib", "ITypeInfo",
              function(x) {
                 lib = .Call("R_GetTypeInfoLibrary", x)	
                 lib@fileName = queryTypeLibPath(lib)

                 lib
         })

setMethod("getTypeLib", "COMIDispatch",
              function(x) {
                  getTypeLib(getTypeInfo(x)[[1]]);
         })



setGeneric("isOptional", function(obj, context = NULL) standardGeneric("isOptional"))
setMethod("isOptional", "ParameterDescription", function(obj, context = NULL) {
    obj@optional
})


setGeneric("getNameIDs", function(obj, context) standardGeneric("getNameIDs"))
setMethod("getNameIDs", c("FunctionDescription", "ITypeInfo"), 
               function(obj, context) {
		  getNameIDs(c(obj@name, names(obj@parameters)), context)
               })

setMethod("getNameIDs", c("ITypeInfo"), 
#
# Loop over the functions and get the ids for its names.
# Return a list with elements for each function.
#
#XXX Doesn't work
#

               function(obj, context) {
                  funcs = getElements(obj)
	          x = lapply(names(funcs),  function(name) getNameIDs(funcs[[name]], obj))
		  names(x) <- names(funcs)
		  x
               })

#setMethod("getNameIDs", c("ITypeLib"), 
#               function(obj, context) {
#	          lapply(names(context), function(name) getNameIDs( context[[name]])
#               })

setMethod("getNameIDs", c("character", "ITypeInfo"), 
               function(obj, context) {
                  if(length(obj) == 0)
                     return(numeric(0))

                  x = .Call("R_GetIDsOfNames", context, obj, FALSE, PACKAGE="SWinTypeLibs")
		  names(x) <- obj
		  x
               })

setMethod("getNameIDs", c("character", "COMIDispatch"), 
               function(obj, context) {
                  if(length(obj) == 0)
                     return(numeric(0))

                  x = .Call("R_GetIDsOfNames", context, obj, TRUE, PACKAGE="SWinTypeLibs")
		  names(x) <- obj
		  x
               })


setMethod("getNameIDs", c("numeric", "ITypeInfo"), 
               function(obj, context) {
	          els = getElements(context)
	          names(els)
	          getNameIDs(names(els)[obj], context)
#Wrong.                  els = .Call("R_GetFuncNames", context, obj, PACKAGE="SWinTypeLibs")
#		  getNameIDs(els, context)
               })


	
# setMethod("getNameIDs", c("FunctionDescription", "ITypeInfo"), 
#               function(obj, context) {
#                 getNameIDs(obj@memid, context)
#               })

setMethod("getNameIDs", c("FunctionDescription", "ITypeInfo"), 
               function(obj, context) {
                 getNameIDs(c(obj@name, names(obj@parameters)), context)
               })



getTypeLibInfo =
function(lib)
{
  if(!is(lib, "ITypeLib"))
    stop("Argument to getTypeLibInfo must be an ITypeLib")

  ans = .Call("R_getLibraryAttr", lib, PACKAGE = "SWinTypeLibs")
  names(ans) <- c("guid", "lcid", "syskind", "major", "minor", "flags")


  ans = new("ITypeLibIdentifier", 
             guid = ans$guid,
             version = c(ans$major, ans$minor),
             lcid = ans$lcid,
             flags = ans$flags,
             syskind = ans$syskind)


  ans
}		


getActiveObject =
function(info)
{
  .Call("R_getActiveObject", info)
}


setGeneric("getTypeName",
   function(x, ...)
      standardGeneric("getTypeName"))

setMethod("getTypeName", "COMIDispatch",
    function(x, ...)
       getTypeName(getTypeInfo(x), ...))


setOldClass(c("ListOfITypeInfo", "list"))

setMethod("getTypeName", "ListOfITypeInfo",
    function(x, ...)
       unname(sapply(x, getTypeName, ...)))

setMethod("getTypeName", "ITypeInfo",
    function(x, ...) {
       getTypeDocumentation(x, 0L)["name"]  # 0 will be mapped to -1L which is MEMBERID_NIL
    })

dcomType = getTypeName

