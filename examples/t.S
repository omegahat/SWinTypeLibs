tt = function(a, b = 2, k = "")
{
 browser()
  print(sys.call()[-1])
  do.call("print", as.list(sys.call()[-1]))
}
