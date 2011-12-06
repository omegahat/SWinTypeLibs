\name{LoadTypeLib}
\alias{LoadTypeLib}
\alias{LoadTypeLib,character-method}
\alias{LoadTypeLib,IContainingTypeLib-method}
\alias{LoadTypeLib,ITypeInfo-method}
\alias{LoadTypeLib,COMIDispatch-method}
\alias{LoadTypeLib,ITypeLibIdentifier-method}
\title{Load Type Library into R}
\description{
 This loads the specified file which is assumed
 to be a type library into R and returns
 a reference or handle to that library object.
 This can then be used in subsequent calls to 
 query the contents of the library.
}
\usage{
LoadTypeLib(fileName)
}
\arguments{
  \item{fileName}{character vector of length 1.
  This is the name of the file to load.}
}
\details{
 This calls  the routine \code{LoadTypeLib()}
 in the Microsoft API.
}
\value{
 An object of class \code{ITypeLib}
 containing a reference to the C-level handle
 for the loaded library.
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{getTypeLibTypes}}
 \code{\link{getTypeDocumentation}}
}

\examples{
 if(file.exists("c:\\\\WINNT\\\\system32\\\\wmpui.dll")) {
   l = LoadTypeLib("c:\\\\WINNT\\\\system32\\\\wmpui.dll")
   names(l)
   l[[1]]
   print(getTypeDocumentation(l, 14))
   print(getTypeLibTypes(l))
 }
}
\keyword{programming}
%\keyword{reflectance}
