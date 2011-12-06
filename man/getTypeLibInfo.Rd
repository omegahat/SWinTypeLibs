\name{getTypeLibInfo}
\alias{getTypeLibInfo}
\title{Returns information about identifying a type library}
\description{
 This function returns information about the attributes of a
 type libray that identify it, such as the UUID, and its
 major and minor versions number and locale information.
 One can get the type library from an instance of an application
 using  \code{\link{LoadTypeLib}} applied to the 
 \code{\link[RDCOMClient]{COMIDispatch-class}}.  Then, one
 can get the information about that type library 
 so that it can be reloaded without the need to 
 use an instance of the application.
 Specifically, one can call \code{\link{LoadRegTypeLib}}
 with the information returned from this function \code{getTypeLibInfo}
 to load the library in a different session or after it has been released.
}
\usage{
getTypeLibInfo(lib)
}
\arguments{
  \item{lib}{the \code{\link{ITypeLib-class}} object whose information
    is to be queried.}
}
\value{
 An object of class \code{\link{ITypeLibIdentifier-class}}.
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{LoadTypeLib}}
 \code{\link{LoadRegTypeLib}}
}
\examples{
 lib <- LoadRegTypeLib("{00020813-0000-0000-C000-000000000046}", c(1, 1))
 getTypeLibInfo(lib)

\dontrun{
 e = COMCreate("Excel.Application")
 lib = LoadTypeLib(e)
 info = getTypeLibInfo(lib)
 e$Quit()
 LoadTypeLib(info)
}
}
\keyword{programming}
\keyword{interface}

