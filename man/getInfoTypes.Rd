\name{getTypeLibTypes}
\alias{getTypeLibTypes}
\title{Types of elements in Type Library}
\description{
  This identifies the types of the different elements
  in the type library, indicating whether they are classes,
  interfaces, enumerations, etc.
}
\usage{
getTypeLibTypes(lib, which = 1:length(lib), byName = TRUE)
}
\arguments{
  \item{lib}{the \code{ITypeLib} object}
  \item{which}{the indices (starting at 1) of the elements in the type
    library to be queried.}
  \item{byName}{a logical value indicating whether to ....(XXX?)}
}

\value{
 An named integer vector giving whose elements
 give the type of corresponding entry in the type library 
 by enumerated value and human-readable name.
 The enumerated values and names are given in \code{TypeKinds}
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{LoadTypeLib}}
 \code{\link{getTypeInfo}}
 \code{\link{getTypeDocumentation}}
}

%\examples{}
\keyword{programming}
%\keyword{reflectance}
