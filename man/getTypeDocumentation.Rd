\name{getTypeDocumentation}
\alias{getTypeDocumentation}
\alias{getTypeDocumentation,ITypeLib-method}
\alias{getTypeDocumentation,ITypeInfo-method}
\title{Documentation for Type Element}
\description{
 This retrieves the documentation for the specified
entry in  the type library (\code{ITypeLib}) or 
\code{ITypeInfo} object.
}
\usage{
getTypeDocumentation(lib, which)
}
\arguments{
  \item{lib}{the \code{ITypeLib} or \code{ITypeInfo} object in
   which to find the documentation element. }
  \item{which}{the index of the entry (1-based) in the type library
  whose documentation is to be retrieved.
    If \code{lib} is an instance of \code{ITypeInfo-class},
    a value of 0 gives the human-readable name of the 
    type.  Values 1, 2, ... give the names of the
     methods/property accessors.
   }
}
\value{
 A character vector of length 3 giving the 
 \item{DocString}{the documentation string for this element.}
 \item{Help Context}{documentation for the context of this element.}
 \item{Help File}{the name of the help file.}
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{LoadTypeLib}}
 \code{\link{getTypeInfo}}
 \code{\link{getTypeLibTypes}}
}
\examples{
if(file.exists("c:\\\\WINNT\\\\system32\\\\wmpui.dll")) {
  l = LoadTypeLib("c:\\\\WINNT\\\\system32\\\\wmpui.dll")

  getTypeDocumentation(l, 0)

  getTypeDocumentation(l, 14)


  m <- matrix("", length(l), 3)
  for(i in 1:length(l)) {
   m[i, ] <- getTypeDocumentation(l, i)
  }
  print(m)

  w = COMCreate("Word.Application", existing = TRUE)
  lib = LoadTypeLib(w)
  getTypeDocumentation(lib[["Documents"]], 0)
  getTypeDocumentation(lib[["Documents"]], 1)
  lib = NULL
}
}
\keyword{programming}
%\keyword{reflectance}
