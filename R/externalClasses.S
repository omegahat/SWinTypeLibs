getExternalClassDefs =
#
# iface is the RCOMTypeLibraryInterfaceDefinition.
#
function(iface)
{
#XXX look in events also.
   classes = 
    lapply(iface@types,
           function(type) {
             k = lapply(list(type@methods, type@propertyGets, type@propertySets),
                           function(els) {
			       tps = lapply(els,
			                     function(def)
					       def@referencedClasses
                                           )

 	                       tps = unlist(tps, recursive = FALSE)
                               if(length(tps)) {
			         isExternal = sapply(tps, is, "ExternalClassDefinition")
                                 tps = tps[isExternal]
                               }
			      tps

                            })
             unlist(k, recursive = FALSE)
          }
    )

    classes = unlist(classes, recursive = FALSE)

    classNames = sapply(classes, function(x) x@className)
    classes[match(unique(classNames), classNames)]
}			


generateExternalClassInterfaces = 
function(classDefs, ...)
{

   libFiles = sapply(classDefs, function(x) x@library@fileName)
   ifaces = lapply(unique(libFiles),
                    function(libName) {
                        defs = classDefs[libFiles == libName]
                        lib = defs[[1]]@library    

                        classNames = sapply(defs, function(x) x@className)
                        enumNames = sapply(defs, function(x) is(x@info, "ITypeInfoEnum"))
                        dispatch = sapply(defs, function(x) is(x@info, "ITypeInfoDispatch"))

                        generateInterface(lib, classNames[dispatch], enums = classNames[enumNames], events = FALSE, ...)
                    })  

    names(ifaces) = unique(libFiles)

    ifaces
}


foo =
function(externalClasses, obj)
{
     # Or sort by library, and generate

     classNames = sapply(externalClasses, function(x) x@className)

     isEnum =  sapply(externalClasses,
                        function(x)  is(x@info, "ITypeInfoEnum"))

     if(any(!isEnum)) {
       obj@types =
            append(obj@types,
                    lapply(obj@externalClasses[!isEnum],
                       function(def) {
                          generateClass(getTypeDocumentation(def@info, 0),
                                         def@library, defaultClass, defaultListClass,
                                         verbose = verbose)
                       }))
     }

     if(any(isEnum))
        obj@enums[classNames[isEnum]] = lapply(obj@externalClasses[isEnum],
                                                 function(def) 
                                                     getElements(def@info))
}