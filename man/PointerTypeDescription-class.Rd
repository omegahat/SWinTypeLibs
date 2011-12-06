\name{PointerTypeDescription-class}
\docType{class}
\alias{PointerTypeDescription-class}
\title{Description of pointer data type in type library}
\description{
 This is an extension of \code{\link{TypeDescription-class}}
for representing pointers or references to basic data types.
This provides a slot for representing the number of indirect
levels to the actual data type, or in other words, the depth
of the pointer.
}
\section{Objects from the Class}{
 Created in C code.
}
\section{Slots}{
  \describe{
    \item{\code{depth}:}{Object of class \code{"integer"} the number
    of indirect levels to the data structure.
     For example, a pointer to a data type (e.g. \code{type *} in C)
     has depth 1. }
    \item{\code{name}:}{the name of the base type to which this is a pointer.}
  }
}
\section{Extends}{
Class \code{"TypeDescription"}, directly.
}
\section{Methods}{
No methods defined with class "PointerTypeDescription" in the signature.
}

\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
 \code{\link{TypeDescription-class}}
}

\examples{

}
\keyword{classes}
