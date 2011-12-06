library(SWinTypeLibs)
library(RDCOMClient)

e = COMCreate("Excel.Application")

lib = LoadTypeLib("C:\\Program Files\\Microsoft Office\\Office\\EXCEL9.OLB")

print(e[["Visible"]])
id = getNameIDs("Visible", lib[["Application"]])
print(id)
.COM(e, "Visible", .dispatch=2, .ids = id)


wkbks = e[["Workbooks"]]
print(wkbks)
id = getNameIDs("Add", lib[["Workbooks"]])
print(id)
.COM(wkbks, "", .dispatch = 1, .ids = id)

id = getNameIDs("Count", lib[["Workbooks"]])
print(id)
.COM(wkbks, "", .dispatch = 2, .ids = id)

id = getNameIDs(lib[["Workbooks"]])[["Open"]]
fileName = gsub("/", "\\\\", system.file("examples", "duncan.xls", package = "RDCOMClient"))
.COM(wkbks, "", fileName, .dispatch = 1, .ids = id[1:2])


	