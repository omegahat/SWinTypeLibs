library(RDCOMClient)

source("WordCode.R")

# setClass("Range", contains = 'COMIDispatch')

app = new("_Application", getCOMInstance("Word.Application"))
app$Documents$Add()
doc = app$Documents[[1]]

doc$Range()


app$Documents[[1]]$AutoHyphenation = !app$Documents[[1]]$AutoHyphenation
