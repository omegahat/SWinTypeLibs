\name{getTypeLibUUIDs}
\alias{getTypeLibUUIDs}
\alias{getTypeLibUUIDs,ITypeLib-method}
\title{Get all of the UUID/GUIDs of the types in the type library.}
\description{
  This function iterates over all the elements in the type library
 and retrieves the GUID for each.  These are labelled by the
 human-readable name of the element.
}
\usage{
getTypeLibUUIDs(lib)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{lib}{an object of class \code{\link{ITypeLib-class}}}
}
\value{
  A character vector with as many elements as there are elements
in the type library. (See \code{\link{length,ITypeLib-method}}.)
 Each value is a UUID, and the corresponding element in the 
 names vector comes from \code{names(lib)}, and gives the
 human-readable form of the identifier for that element of the type library.
 This is a useful way to map from GUID to type library element.
 This is needed, for example, when matching connection points
 for registering for events to particular interfaces. The connection
 point gives us the UUID and we map this to an interface using
 this function.
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{names,ITypeLib-method}}
 \code{\link{[[,ITypeLib-method}}
}
\examples{
\dontrun{
  library(RDCOMClient)
  e = COMCreate("Excel.Application")
  lib = LoadTypeLibrary(e)
  getTypeLibUUIDs(lib)
}
}
\keyword{interface}
\keyword{programming}

