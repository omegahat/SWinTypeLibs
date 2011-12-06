# This probably isn't used at all. 
	
	
StandardDollarDef =' 
	     idx = match(name, names(GetProperty)) 
             if(!is.na(idx)) {
                return(GetProperty[[name]]())
             }

             if(StrictMethodNameExpansion) m = match
             idx = m(name, names(Methods))
             if(is.na(idx))
                stop("no such method ", name, " for Excel.Application")

              name = names(Methods)[idx] 
              Methods[[idx]]'

StandardDollarGetsDef =' 
	     idx = pmatch(name, names(GetProperty)) 
             if(is.na(idx)) {
                stop("No such method for <class>")
             }
	     GetProperty[[idx]](value)
'


rewriteDollar =
function(def, className)
{
 ids = createTypeVarName(className, c("GetProperty", "Methods"))
 def = gsub("GetProperty", ids["GetProperty"], def)
 def = gsub("Methods", ids["Methods"], def)

 def
}


  # Comes from writeCode method for RCOMTypeInterfaceDefinition
if(FALSE) {
             cat(ids["$"], " =  CompiledCOMAccessor(function(x, name) {", "\n\n", file = file)
             cat(rewriteDollar(StandardDollarDef, def@className[1]), "\n", file = file)
             cat("})\n", file = file)

             cat("setMethod('$', '", def@className[1],"', ", 
                   "get(", ids["$"], ", mode = 'function')", ")\n", file = file, sep = "")  

             cat(ids["$<-"], " =  CompiledCOMAccessor(function(x, name) {", "\n\n", file = file)
             cat(rewriteDollar(StandardDollarGetsDef, def@className[1]), "\n", file = file)
             cat("})\n", file = file)

             cat("setMethod('$<-', '", def@className[1],"', ", 
                   "get(", ids["$<-"], ", mode = 'function')", ")\n", file = file, sep = "") 
}	
