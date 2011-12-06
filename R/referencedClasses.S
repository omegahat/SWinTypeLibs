getReferencedClasses =
function(iface, lib = NULL,  
         defaultClass = "CompiledCOMIDispatch",
         defaultListClass = c("COMList", "COMTypedList"),
	 processExternalClasses = FALSE)
{
  classes = list()
  lapply(iface@types, 	
           function(x) {
	      for(i in list(x@methods, x@propertyGets, x@propertySets)) {
                tmp = lapply(i,
                              function(z) z@referencedClasses)
                classes <<- c(classes, unlist(tmp, recursive = FALSE))
              }
           })

  nn = sapply(classes, is.null)
  classes = classes[!nn]

   # avoid duplicates, within this list and also across the regular types
   # that we were asked to process and the enumerations.
  names(classes) = sapply(classes, function(x) x@className)
  classes = classes[ !duplicated(names(classes)) ]
  classes = classes[ !(names(classes) %in% c(names(iface@types), names(iface@enums))) ]


  if(!is.null(lib)) 
    classes = lapply(classes, recomputeClassDef, names(lib), defaultListClass, defaultClass, lib)

  classes
}			

recomputeClassDef =
 function(klass, libNames, defaultListClass, defaultClass, lib) {
      if(is(klass, "NamedClassDefinition")) {
         className = klass@className
         if(! (className %in% libNames))
            return(klass)
         funcs = getElements(lib[[className]])
         ans = getListClassName(funcs,  className, lib[[className]], lib, defaultListClass, defaultClass)
         if(is.null(ans))
            klass
         else
             ans
	 } else
        klass
}

