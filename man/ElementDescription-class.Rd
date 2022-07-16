\name{ElementDescription-class}
\docType{class}
\alias{ElementDescription-class}
\title{Base class for describing elements of a type library.}
\description{ 
Base class for describing elements such as variables, functions, etc.
within a \code{\link{ITypeInfo-class}} object.
}
\section{Objects from the Class}{A virtual Class: No objects may be created from it.}
\section{Slots}{
  \describe{
    \item{\code{name}:}{Object of class \code{"character"} name of the element being described }
    \item{\code{hidden}:}{Object of class \code{"logical"} indicates whether the element is intended to be 
  hidden or not}
    \item{\code{memid}:}{ for a dispatch type, this is the same as the DISPID value.
    If I recall, this can be used to avoid the lookup in the C code of the
    element before doing the dynamic call.}
  }
}

\section{Methods}{
No methods defined with class "ElementDescription" in the signature.
}

\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
 \code{\link{getElements}}
 \code{\link{FunctionDescription-class}}
 \code{\link{ParameterDescription-class}}
 \code{\link{VariableDescription-class}}
 \code{\link{PointerTypeDescription-class}}
}

%\examples{}
\keyword{classes}
