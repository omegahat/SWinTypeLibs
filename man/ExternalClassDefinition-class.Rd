\name{ExternalClassDefinition-class}
\docType{class}
\alias{ExternalClassDefinition-class}

\title{ExternalClassDefinition for representing referenced DCOM types in other libraries.}
\description{ 
 This class is used when we encounter a reference to a DCOM type
 (i.e. of class \code{\link{TypeDescriptionRef-class}})
 in a method's parameters or return value description,
 and that referenced class is actually defined in another type library
 than the one in which the method is located.
 As we process the original type library, we collect these references
 to external types using this class and then process them 
 separately at the end.
 
}
\section{Objects from the Class}{
Objects can be created by calls of the form \code{new("ExternalClassDefinition", ...)}.
These contain information about the name of the DCOM type
(computed via \code{\link{getRefTypeName}} (with the \code{follow} argument being
\code{TRUE})), and the information about the library in which the
class is defined and the DCOM type (i.e. the \code{\link{ITypeInfo-class}} object)
in that library for that type.
This information makes this object fully self-describing and allows
us to generate an interface to that type.
}
\section{Slots}{
  \describe{
    \item{\code{info}:}{Object of class \code{"ITypeInfo"}. 
     This is the resolved \code{\link{ITypeInfo-class}} that 
     was computed when resolving the reference to the class
     from the \code{\link{TypeDescriptionRef-class}} information.
     This can then be used in subsequent processing to generate
     an interface to this DCOM type.
    }
    \item{\code{library}:}{Object of class \code{"ITypeLib"}.
     a reference to the type library in which the class is actually
     defined and in which the value of the \code{info} slot
     is located.}


    \item{\code{code}:}{Object of class \code{"character"} inherited from
     \code{\link{ClassDefinition-class}}.}
    \item{\code{className}:}{Object of class \code{"character"}  inherited from
     \code{\link{ClassDefinition-class}}. }
  }
}
\section{Extends}{
Class \code{"NamedClassDefinition"}, directly.
Class \code{"ClassDefinition"}, by class \code{"NamedClassDefinition"}.
This extends \code{NamedClassDefinition} merely to inherit
the show method so that it prints correctly.
}
\section{Methods}{
No public methods defined with class "ExternalClassDefinition" in the signature.

}

\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{ClassDefinition-class}}
 \code{\link{NamedClassDefinition-class}}
 \code{\link{InlineClassDefinition-class}}

 \code{\link{generateClass}}
 \code{\link{generateInterface}}
}

%\examples{}
\keyword{classes}
