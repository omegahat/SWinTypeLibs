\name{getTypeLibEventElements}
\alias{getTypeLibEventElements}
\title{Identify all event interfaces in a type library.}
\description{
 This function implements a heuristic approach to 
 finding all the  elements within a type library
 that are interfaces used to handle events
 from a DCOM event source.

 There are other approaches that may work better
 for different type libraries that use different
 conventions for identifying the event interface elements.
 Indeed, there are two versions of this function in
 the \code{generate.S} file in the source of this package.
 By default, this function is invoked if the caller 
 of the top-level function \code{\link{generateInterface}}
 does not supply the names of the event interfaces.
 One can use a different function and supply the element names
 in the type library returned from that as the value
 of the \code{events} argument for \code{\link{generateInterface}}.
 So in this sense, this is just the default function that
 one can substitute with more customized approaches.

 This function looks for all elements named Events.
 It then removes any elements named I*
}
\usage{
getTypeLibEventElements(lib)
}
\arguments{
  \item{lib}{the type library in which to look for event interface elements.}
}
\value{
 A character vector giving the names of the elements in the type library
 that are event interfaces.

 This could usefully return the elements in the type library directly,
 rather than just their names.
}

\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
 \code{\link{generateInterface}}
 \code{\link{generateEventInterface}}
}
\examples{
\dontrun{
  # Load excel type library. Use minor version of 1 to get anything higher than that.
  lib = LoadRegTypeLib("{00020813-0000-0000-C000-000000000046}", c(1, 1))
  getTypeLibEventElements(lib)
}
}
\keyword{programming}
\keyword{interface}
\concept{meta-programming}

