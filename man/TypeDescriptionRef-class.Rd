\name{TypeDescriptionRef-class}
\docType{class}
\alias{TypeDescriptionRef-class}
\title{Reference to a type description object}
\description{
 In the type library, some elements will provide
 type information that refers to other classes or interfaces
 in the same library. This is then given as a reference or a path to the
 particular type description. 
 This class represent such a reference. It holds the handle or identifier
 which we can use to resolve the actual type description.
}
\section{Objects from the Class}{
Created internally in the C code.
}
\section{Slots}{
  \describe{
    \item{\code{reftype}:}{Object of class \code{"numeric"} the identifier in the type library
     which one uses to resolve the actual type to which this one refers.}

   \item{\code{name}:}{inherited from \code{\link{TypeDescription-class}}}
  }
}
\section{Extends}{
Class \code{"TypeDescription"}, directly.
}
\section{Methods}{
  \describe{
    \item{getTypeInfo}{\code{signature(obj = "ITypeInfo", which = "TypeDescriptionRef")}: resolve
    the actual type description. }
  }
}

\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{

}


\examples{

}
\keyword{classes}
