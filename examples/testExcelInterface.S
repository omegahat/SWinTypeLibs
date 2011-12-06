library(SWinTypeLibs)

library(RDCOMClient)
library(RDCOMServer)

library(RDCOMEvents)

source("ExcelCode.S")
load("EventTable.rda")
	
e = new("_Application", COMCreate("Excel.Application"))

if(e$Ready)
  e[["Visible"]] = TRUE

books = e$Workbooks
books$Add()	
book = books[[1]]

ss = WorkbookEventsHandler(list(NewSheet = function(sheet) {
                                            cat("new Excel sheet\n")
                                           },
                                SheetActivate = function(sheet) {
                                                  cat(" switched to sheet", sheet[["Name"]], "\n")
                                                }
                               ))

connectEventHandlers(book, ss)
