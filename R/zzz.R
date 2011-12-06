.First.lib <-
function(lib, pkg) {
 library.dynam("SWinTypeLibs", pkg, lib)
# library(methods)
}


.onLoad <-
function(lib, pkg) {
 library(methods)
 library.dynam("SWinTypeLibs", pkg, lib)
}
