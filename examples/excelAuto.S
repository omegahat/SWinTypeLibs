# Load the source library.
source("excel.S")

setClass("_Application", contains = "CompiledCOMIDispatch")

tmp = generateClass("_Application", lib)
writeCode(tmp, "ApplicationEx.S")

source("ApplicationEx.S")

library(RDCOMClient)
e = new("_Application", COMCreate("Excel.Application")) 


e$Ready

