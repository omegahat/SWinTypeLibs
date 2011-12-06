\name{LoadRegTypeLib}
\alias{LoadRegTypeLib}
\title{Load type library given UUID and version information}
\description{
 Like \code{\link{LoadTypeLib}}, this function 
 loads a type library into R and returns a reference/handle
 to it.  However, how it finds the type library is different.
 \code{\link{LoadTypeLib}} is given the fully qualified
 location of the file that contains the type library.
 This function however consults the Window's registry
 to find the type library file and then load its.
 It uses the UUID that  identifies the 
 type library, along with version information
 and the locale identifier to determine which specific type
 library to load.
}
\usage{
LoadRegTypeLib(guid, version, lcid = 0)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{guid}{the UUID . Currently, this is a string/character vector
   rather than an object of class UUID.}
  \item{version}{an integer vector giving the major and minor version
    number in that order.}
  \item{lcid}{a numeric/real value identifying the locale for the type library.}
}
\details{
 This calls the C-level routine \code{LoadRegTypeLib}.
 It consults the registry. We do not.
}
\value{
 An object of class \code{\link{ITypeLib-class}}.
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{LoadTypeLib}}
 \code{\link{getTypeLib}}
 \code{getTypeLibInfo} (unexported from SWinTypeLibs)
  retrieves the UUID, version and locale information from
  a \code{IContainingTypeLib} object derived from a \code{ITypeInfo}.
}
\examples{
  # Get the excel type library.
  try(lib <- LoadRegTypeLib("{00020813-0000-0000-C000-000000000046}", c(1, 1)))
}
\keyword{programming}

