comAccessors =
function(info, className, where = NULL, useFunctions = FALSE) {

  get <- function(x, name) {
    name <- tolower(name)
    if(!useFunctions && !is.na(match(name, info$functions))) {
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

  if(useFunctions) {
   for(i in info$functions) {
     def <- info$definitions[i]
     createFunctionDefinition(i, def)
     setMethod(i, className, f)
   }
  }

  setMethod("$", className, get, where = where)
  setMethod("$<-", className, set, where = where)
}
	
createFunctionDefinition <-
#
# Need to get default arguments
#
# Process the arguments in the call and drop the ones that are missing
# if they can be missing in the COM function. 
#
# Want to build the function in S, not strings
#
function(def, name = NULL)
{
 if(is.null(name) || name == "")
  name <- def@name

 f <- function(){}

 for(i in names(def@parameters)) {
  formals(f)[i] <- alist(xx=)
 }

 # Now the body.
 b = expression()
 for() {
 }

 environment(f) <- globalenv()
 return(f)

if(FALSE) {
 str <- paste("function(", 
              paste(c(".obj", names(def@parameters)), collapse=", "), 
              ") {\n",
              " .COM(.obj, \"", name, "\", ", 
              paste(names(def@parameters), collapse=", "),
              ")\n}",
              sep="")
 print(str)
 eval(parse(text = str))
}
}

comFunctionAccessor <-
function(info, className, where = NULL)
{

}
