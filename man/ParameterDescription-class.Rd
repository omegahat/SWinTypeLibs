\name{ParameterDescription-class}
\docType{class}
\alias{ParameterDescription-class}
\title{Description of a parameter of a function definition}
\description{
 This describes an individual parameter in a function definition
within a type library.
 It provides the name of the parameter, its type
and default value,  and its style (e.g. in or out, whether it is optional).
}
\section{Objects from the Class}{
Objects can be created by calls of the form \code{new("ParameterDescription", ...)}.
}
\section{Slots}{
  \describe{
    \item{\code{name}:}{Object of class \code{"character"} 
        the name of the parameter, used in providing arguments in calls by name.}
    \item{\code{type}:}{Object of class \code{"TypeDescription"} 
      the information for this parameter describing the 
      data type that is expected in the call to this function for this parameter }
     \item{\code{style}:}{Object of class \code{"ParameterStyle"} 
      the call style for this parameter, indicating whether it is an
      in, out or in/out argument by which it transfers content to and from
      the function. }
     \item{\code{defaultValue}:}{Object of class \code{"ANY"} 
      an S object giving the default value for this argument if any.
   }
  }
}

\section{Methods}{
No methods defined with class "ParameterDescription" in the signature.
}

\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
 \code{\link{getElements}}
 \code{\link{ElementDescription-class}}
 \code{\link{FunctionDescription-class}}
 \code{\link{VariableDescription-class}}
 \code{\link{TypeDescription-class}}
 \code{\link{PointerTypeDescription-class}}
}

\examples{

}
\keyword{classes}
