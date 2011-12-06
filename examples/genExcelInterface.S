# Get the type library by creating an instance of the application.
# This avoids having to know where the type library is located.

library(SWinTypeLibs)

library(RDCOMClient)
library(RDCOMServer)

library(RDCOMEvents)

	
e = COMCreate("Excel.Application")
options(error=recover)	

lib = LoadTypeLib(e)	

processEvents = TRUE

cat("Generating interface\n")
	
# Generate the code to interface to ALL the classes.
#tmp = generateInterface(lib, "Workbooks", events = FALSE)
#gctorture(TRUE)

	
tmp = generateInterface(lib,
	"ShapeNodes",
#                        c("_Application"), # "Workbooks", "Workbook", "_Workbook", "Range"),
                        events = processEvents,
	                verbose = TRUE)

cat("Finished generating interface\n")


out = writeCode(tmp, "ExcelCode.S") # "tmp"

if(TRUE) {
# Now load this into our session. 
# Since we put it into a single file (rather than a directory)
# we need only source that one large file. It will contain
# the runtime.S and common.S code.

cat("Reading the source code back into R\n")

source("ExcelCode.S")

if(processEvents)
  load("EventTable.rda")
}



