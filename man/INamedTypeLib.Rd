\name{INamedTypeLib}
\alias{INamedTypeLib}
\title{Create instance of INamedTypeLib class from regular ITypeLib}
\description{
 This is a constructor function to create an instance of
 the \code{\link{INamedTypeLib-class}} which 
 maintains a cache of its components: names, UUIDs and \code{\link{ITypeInfo-class}}
 elements.  It takes a regular \code{\link{ITypeLib-class}} 
 object and creates the cached version  that can be used
 in its place.
 The cache avoids recomputing the same elements repeatedly.
}
\usage{
INamedTypeLib(lib)
}
\arguments{
  \item{lib}{an object of class \code{\link{ITypeLib-class}}.}
}
\value{
 An instance of the class \code{\link{INamedTypeLib-class}}.
}

\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
 \code{\link{ITypeLib-class}}
 \code{\link{LoadTypeLib}}
 This is used in \code{\link{generateInterface}}
 as we make intensive use of the library's contents
 and so caching improves performance.
}
%\examples{}
\keyword{interface}
\keyword{programming}


