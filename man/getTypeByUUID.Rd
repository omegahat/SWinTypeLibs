\name{getTypeByUUID}
\alias{getTypeByUUID}
\title{Finds element in ITypeLibrary by UUID}
\description{
 This function takes a library and a UUID and finds the 
 element by comparing the UUID to those of all elements
 in the library.
 This function is used in the \code{[[} method for 
 indexing the type library by name/uuid.
}
\usage{
getTypeByUUID(uuid, lib)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{uuid}{the UUID string (with enclosing braces (\{\}).}
  \item{lib}{the \code{\link{ITypeLib-class}} object representing the type library.}
}
\value{
  An object of class \code{\link{ITypeInfo-class}}
}

\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{names,ITypeLib-method}}
 \code{\link{[[,ITypeLib,character-method}}
}
%\examples{}
\keyword{programming}
