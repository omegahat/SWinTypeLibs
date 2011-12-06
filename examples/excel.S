library(SWinTypeLibs)

files = c("C:/Microsoft Office/OFFICE11/Excel.exe", 
          "C:\\Program Files\\Microsoft Office\\Office\\EXCEL9.OLB")

for(i in files) {
 if(file.exists(i)) {
  cat("Reading", i, "\n")
  lib = LoadTypeLib(i)
  break
 }
}

	
