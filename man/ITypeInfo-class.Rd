\name{ITypeInfo-class}
\docType{class}
\alias{ITypeInfo-class}
\alias{ITypeInfoCoClass-class}
\alias{ITypeInfoModule-class}
\alias{ITypeInfoDispatch-class}
\alias{ITypeInfoEnum-class}
\alias{ITypeInfoInterface-class}
\alias{ITypeInfoRecord-class}
\alias{ITypeInfoUnion-class}
\alias{ITypeInfoAlias-class}
\alias{[[,ITypeInfoCoClass,numeric-method}
\alias{[[,ITypeInfoCoClass,ANY-method}
\alias{[,ITypeInfoCoClass-method}
\alias{[,ITypeInfoCoClass,ANY-method}

\title{Information about an interface, class, enumeration, alias, etc.}
\description{
 These classes provide a handle to C-level data structures
 that describe an element within a type library.
 From an instance we can obtain information about 
 the methods and variables within a class, 
 values in an enumeration, aliases or type definitions 
 for an element in terms of another, and so on.

 \code{ITypeLibEntry} is a simple extension that provides
 information about the origin of the information.  This indicates that
it was obtained from a type library rather than a DCOM object.

}
\section{Objects from the Class}{

Objects can be created by calls of the form \code{\link{getTypeInfo}}
and other covenience operators such as \code{lib[["Application"]]}
where \code{lib} is a \code{\link{ITypeLib-class}} object.

}
\section{Slots}{
  \describe{
    \item{\code{ref}:}{Object of class \code{"externalptr"} 
      the reference  to the C-level \code{ITypeInfo} value.
    }
    \item{\code{type}:}{Object of class \code{"integer"} 
    a named integer that indicates what type of element this information
    object refers to. This is an enumeration, dispatch, coclass, alias, etc.
    The name provides a human-readable form; the integer value is the
    enumeration value from C.}
    \item{\code{guid}:}{Object of class \code{"character"}.
    This is the stringified version of the UUID that uniquely identifies
    this element/interface.}
  }
}

\section{Methods}{
  \describe{
    \item{getEnum}{\code{signature(x = "ITypeInfo")}: get a list of the enumeration elements within 
      this information object.}
    \item{getFuncs}{\code{signature(x = "ITypeInfo")}: 
       get a list of all of the function objects within this information object.}
    \item{getTypeDocumentation}{\code{signature(lib = "ITypeInfo")}: 
     get the name and documentation string for the information object.}
    \item{getVars}{\code{signature(x = "ITypeInfo")}: get the variables within the information object.}
  }
}

\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
\code{\link{getTypeInfo}}
\code{\link{LoadTypeLib}}
}

%\examples{}
\keyword{classes}
