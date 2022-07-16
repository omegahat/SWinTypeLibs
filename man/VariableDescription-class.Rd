\name{VariableDescription-class}
\docType{class}
\alias{VariableDescription-class}
\title{Describes a variable in an ITypeInfo object}
\description{
This class is used to provide an S-level description
of a variable in a type library.
}
\section{Objects from the Class}{
 Generated in C code via \code{\link{getElements}}.
}
\section{Slots}{
  \describe{
    \item{\code{type}:}{Object of class \code{\link{TypeDescription-class}} 
       description of the type of the variable}

   \item{\code{name}:}{inherited from \code{link{ElementDescription-class}}}
   \item{\code{memid}:}{inherited from \code{link{ElementDescription-class}}}
   \item{\code{hidden}:}{inherited from \code{link{ElementDescription-class}}}    
 }
}
\section{Extends}{
Class \code{"ElementDescription"}, directly.
}
\section{Methods}{
No methods defined with class "VariableDescription" in the signature.
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
 \code{\link{TypeDescription-class}}
 \code{\link{PointerTypeDescription-class}}
}

%\examples{}
\keyword{classes}
