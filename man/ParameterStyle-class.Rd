\name{ParameterStyle-class}
\docType{class}
\alias{ParameterStyle-class}
\title{Qualifer of a parameter description}
\description{
 This provides additional information about a parameter in a COM method
in terms of how it is expected to be used.
It provides information about whether it is an input or an output
variable.
}
\section{Objects from the Class}{
 Create in C code.
}
\section{Slots}{
  \describe{
    \item{\code{In}:}{Object of class \code{"logical"} is it an input variable }
    \item{\code{out}:}{Object of class \code{"logical"} is it an output variable, i.e. containing a return value }
    \item{\code{optional}:}{Object of class \code{"logical"} is the parameter optional in a call, or in other words, can the caller omit this argument in a call.}
    \item{\code{lcid}:}{Object of class \code{"logical"} is the locale information of the client application }
    \item{\code{retval}:}{Object of class \code{"logical"} is the return value of the method.}
  }
}

\section{Methods}{
No methods defined with class "ParameterStyle" in the signature.
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{getElements}}
 \code{\link{ParameterDescription-class}}
}

\examples{

}
\keyword{classes}
