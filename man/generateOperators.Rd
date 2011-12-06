\name{generateOperators}
\alias{generateOperators}
\title{Define \$ methods for typed COM class}
\description{
 This is used by \code{\link{createCOMSClass}}
 to create the \code{$} and \code{$<-}
 methods for the new S4 class that mirrors the
 COM class. This function defines S4 methods
 for these operators that contain information about the
 available properties and methods for the COM object
 and so provide structured access to them.
}
\usage{
generateOperators(libEntry, className)
}
\arguments{
  \item{libEntry}{the \code{\link{ITypeInfo-class}} object
   from which to get the property and method descriptions.} 
  \item{className}{the name of the S4 class for which to define the methods}
}
\value{
A list giving the two functions that were registered
as get and set methods for \code{$} and \code{$<-} respectively.
}

\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{createCOMSClass}}
 \code{\link{computeFunctionInformation}}
}
\examples{

}
\keyword{programming}
\keyword{interface}



