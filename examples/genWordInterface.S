library(SWinTypeLibs)
library(RDCOMClient)
library(RDCOMServer)
library(RDCOMEvents)	

app = COMCreate("Word.Application")
lib = LoadTypeLib(app)

interface=generateInterface(lib, c("_Application","_Document","Documents", "Range"), events=FALSE)

writeCode(interface, "WordCode.R")

#source("WordCode.R")


