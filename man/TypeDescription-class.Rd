\name{TypeDescription-class}
\docType{class}
\alias{TypeDescription-class}
\title{Description of a data type from a type library.}
\description{
This provides an S-level description of a 
data type within a type library.
}
\section{Objects from the Class}{
Objects can be created by calls of the form \code{new("TypeDescription", ...)}.

}
\section{Slots}{
  \describe{
    \item{\code{name}:}{Object of class \code{"character"} the name of the data type being described}
  }
}

\section{Methods}{
No methods defined with class "TypeDescription" in the signature.
}

\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{ 
 \code{\link{getElements}}
 \code{\link{ElementDescription-class}}
 \code{\link{FunctionDescription-class}}
 \code{\link{ParameterDescription-class}}
 \code{\link{VariableDescription-class}}
 \code{\link{PointerTypeDescription-class}}
}

%\examples{}
\keyword{classes}
