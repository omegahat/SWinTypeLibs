
source("excel.S")

iface = generateInterface(lib, c("_Application", "_Workbook", "_Worksheet",
                                 "Workbooks", "Worksheets"))


writeCode(iface, tempdir())