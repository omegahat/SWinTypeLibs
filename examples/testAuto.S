.Call <-
function(name, ...)
{
  cat("Started .Call", name, "\n")
  ans = get(".Call", "package:base")(name, ...)
  cat("Finished .Call", name, "\n")
  ans
}
rm(.Call)
options(error=recover)
library(SWinTypeLibs)
source("excel.S")
source("autoInterface.S")
.COMSClassMap = character()
library(RDCOMClient)
createCOMSClass(lib[["Workbooks"]], "Workbooks")
cat("Creating application class\n")
createCOMSClass(lib[["_Application"]], "ExcelApplication")
save.image()

if(!exists("no")) {
lapply(names(lib), function(x) { cat(x,"\n"); 
     if(x %in% c("ColorFormat", "ShapeNode", "Constants"))
        return(NULL)
             createCOMSClass(lib[[x]], x)
})
}


