\name{FunctionDescription-class}
\docType{class}
\alias{FunctionDescription-class}
\alias{FunctionInvokeDescription-class}
\alias{PropertyGetDescription-class}
\alias{PropertySetDescription-class}
\alias{PropertySetRefDescription-class}
\title{Description of a function from a Type Library.}
\description{
 These classes describe a function element in a type library.
The different types of functions refer to regular 
functions (\code{FunctonInvokeDescription})
and property accessors.
The different classes are used simply for type information
and dispatching rather than additional slots.
The description object gives details about 
the parameters, the return type, the type of function (regular or property accessor).

One can determine from the collection of functions whether a property is read-only
by determining if there is no set method for that property.
}
\section{Objects from the Class}{
\code{\link{getFuncs}}
}
\section{Slots}{
  \describe{
    \item{\code{returnType}:}{Object of class \code{"TypeDescription"} 
       a description of the return values for this function.}
    \item{\code{parameters}:}{Object of class \code{"list"}
        a list of \code{\link{ParameterDescription-class}}
        objects detailing the different parameters for the function.
     }
    \item{\code{invokeType}:}{Object of class \code{"integer"} 
     a named integer indicating how this function should be invoked,
     i.e. as a propertyget, propertyput, function call,
      propertyrefput,
     }
    \item{\code{name}:}{Object of class \code{"character", from class "ElementDescription"} ~~ }
    \item{\code{hidden}:}{Object of class \code{"logical", from class "ElementDescription"} ~~ }
  }
}
\section{Extends}{
Class \code{"ElementDescription"}, directly.
}
\section{Methods}{
No methods defined with class "FunctionDescription" in the signature.
}

\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{ElementDescription-class}}
 \code{\link{getFuncs}}
}

%\examples{}
\keyword{classes}
