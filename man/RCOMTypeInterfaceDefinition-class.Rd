\name{RCOMTypeInterfaceDefinition-class}
\docType{class}
\alias{RCOMTypeInterfaceDefinition-class}
\alias{RCOMEventInterfaceDefinition-class}
\alias{RCOMDispatchInterfaceDefinition-class}

\title{Classes for representing R code to interface to DCOM classes.}
\description{
These classes provide basic structures for representing
R code that interfaces to (i) a DCOM interface/class,
and (ii) a DCOM event handler interface.

\code{RCOMDispatchInterfaceDefinition}  is used
to represent code that interfaces to the methods and
property accessors of a DCOM interface. The elements
define the R functions that interface to these DCOM methods.

The \code{RCOMEventInterfaceDefinition} class does something
similar but provides code to create an event handler
for the particular DCOM event interface.
It provides a constructor function for this, along with
template functions to implement the event server methods.
}
\section{Objects from the Class}{
Objects can be created by calls of the form 
\code{new("RCOMDispatchInterfaceDefinition", ...)}
and \code{new("RCOMEventInterfaceDefinition", ...)}.
However, they are not expected to be used except by 
developers. They are exported so that others can
use their contents.
Instead, they are created in calls to 
\code{\link{generateClass}} and \code{\link{generateEventInterface}}.
}
\section{Slots}{
  \describe{
    \item{\code{className}:}{Object of class \code{\link{ClassDefinition-class}},
       code defining an S4 class representing the interface.}
    \item{\code{methods}:}{Object of class \code{"list"},
      a list of \code{\class{COMMethodDefinition-class}} elements
      corresponding to the methods (not property accessors)
      in the DCOM interface.}
    \item{\code{guid}:}{Object of class \code{"character"},
     a collection of the UUIDs implemented by the interface.
     In general, there will be just one of these.
     This allows us to map between the human-readable name of the 
     DCOM interface and its precise, unique identifier.}
    \item{\code{interfaceName}:}{Object of class \code{"character"},
     the name of the DCOM interface/class for which the code is being
     generated. This may be different from the R class name
     used to interface to the DCOM class.}

 
    
    \item{\code{constructor:}}{a character string giving the 
     code for an R function that creates an R event handler
     for this type of DCOM interface.
     See \code{\link[RDCOMEvents]{createEventServerFromTemplate}}.}
    \item{\code{template:}}{an object of class
     \code{\link[RDCOMEvents]{COMEventServerInfo-class}}
      that is used to provide a stub or template for each 
     method in the DCOM interface along with information 
     about the parameters and their types, and the return value of the
     method.}

    \item{\code{propertyGets:}}{a list of 
     \code{\link{COMPropertyGetDefinition-class}} objects
     corresponding to the property read-accessors methods in the DCOM
     interface.}
    \item{\code{propertySets:}}{a list of 
     \code{\link{COMPropertySetDefinition-class}} objects
     corresponding to the property write-accessors, i.e. the methods that set the 
     value of a property, in the DCOM
     interface.}

  }
}
\section{Methods}{
  \describe{
    \item{writeCode}{\code{signature(def = "RCOMTypeInterfaceDefinition", file = "character")}: ... }
    \item{writeCode}{\code{signature(def = "RCOMTypeInterfaceDefinition", file = "ANY")}: ... }
  }
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{generateInterface}}
 \code{\link{writeCode}}

 Non-exported methods in \code{generate.S}:
 \code{generateClass}, \code{generateMethod}
}
\examples{
}
\keyword{classes}
