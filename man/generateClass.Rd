\name{generateClass}
\alias{generateClass}
\title{Create R code to interface to DCOM class/type}
\description{
 This function is part of the code generation 
 mechanism to interface to DCOM classes/types in
 a type library. This function generates
 the R code to access the methods and properties
 of the DCOM interface along with additional information
 about class definitions, etc.
}
\usage{
generateClass(id, lib, defaultClass = "CompiledCOMIDispatch", 
              defaultListClass = c("COMList", "COMTypedList"),
              verbose = FALSE)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{id}{a string giving the human-readable name of the DCOM interface
   in the type library.}
  \item{lib}{the type library of class \code{\link{ITypeLib-class}}
   in which the interface information is to be found.}
  \item{defaultClass}{a string giving the name of the R class
    that is to be used as the base class for new classes
    defined  to represent DCOM classes. This is used
    if the DCOM class is not a list. See 
    \code{computeClassName} and \code{getListClassName}
    in the file \code{generate.S}.
   }
  \item{defaultListClass}{either a function
   or a character vector. If this is a function,
   this is called to determine the class definition
   of the new S class to be constructed to represent
   references to the DCOM objects of this type.
   The function must return an object derived
   from \code{\link{ClassDefinition-class}}.
   This allows the caller to specify her own function
   to compute S class definitions from DCOM types.
    
   Alternatively, if this is given as a character vector,
   it should give the names of classes that are used
   for "untyped" and "typed" lists, in that order.
   By typed lists, we mean where we know the (base) type of the
   elements in the DCOM list.  If these are general
   VARIANT or IDispatch objects, then it is an untyped list.
  }
  \item{verbose}{a logical value, with \code{TRUE} indicating that
    output reporting progress and the current activity should be displayed
    on the console.}
}
\value{
 An object of class
 \code{\link{RCOMDispatchInterfaceDefinition-class}}
}

\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
 \code{\link{generateInterface}}
 \code{\link{generateEventInterface}}
}

\examples{

}
\keyword{programming}
\keyword{interface}
\concept{meta-programming}


