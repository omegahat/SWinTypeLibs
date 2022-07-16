\name{COMOperationDefinition-class}
\docType{class}
\alias{COMOperationDefinition-class}
\alias{COMMethodDefinition-class}
\alias{COMEventDefinition-class}
\alias{COMPropertySetDefinition-class}
\alias{COMPropertyGetDefinition-class}
\alias{show,COMOperationDefinition-method}

\title{Classes giving meta-information about operations of a DCOM interface}
\description{
 These classes provide a means to represent a
 textual version of a function that
 provides an interface from R to the elements
 or operations described in a type library about a particular
 interface. These elements describe methods, properties and the
 accessor methods to get and set their values, and 
 methods for event handler interfaces.
 And these classes represent the R code (as a character vector)
 that provides the interface to the underlying methods.

 In addition to the text of the function,
 these objects also carry with them information about the
 classes that they reference. This  information
 is used when serializing the code for future use
 to define new classes that arise at run-time
 but which are not explicitly processed in the type
 library.
}
\section{Objects from the Class}{
Objects can be created by calls of the form 
  \code{new("COMMethodDefinition", ...)}, etc..

\code{COMOperationDefinition} is a virtual class.
}
\section{Slots}{
  \describe{
    \item{\code{.Data}:}{Object of class \code{"character"},
      this is where the text giving the R function definition
      to interface to the DCOM method is stored.
      }
    \item{\code{referencedClasses}:}{Object of class \code{"list"},
      this is a list of \code{\link{ClassDefinition-class}} objects.
      Each element provides a definition of a class that was referenced
      in the DCOM method that might need to be defined.
      These arise from the parameters and the return type of the
      DCOM method.}
  }
}
\section{Extends}{
Class \code{"character"}, directly, with explicit coerce.
Class \code{"vector"}, by class "character", with explicit coerce.
}
\section{Methods}{
  \describe{
    \item{show}{\code{signature(object = "COMOperationDefinition")}: 
      this provides a method to display the code in this object
      so that it is easy to read. }
  }
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.net/SWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}
\seealso{
 \code{\link{generateInterface}}
}
%\examples{}
\keyword{classes}
\keyword{programming}
\concept{meta-programming}
