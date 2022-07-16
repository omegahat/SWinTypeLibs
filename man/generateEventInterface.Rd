\name{generateEventInterface}
\alias{generateEventInterface}
\title{Create R interface to DCOM event class}
\description{
 This function creates R code to provide an interface
 to a DCOM class or type that defines an event handler,
 i.e. a DCOM interface that is used to respond to events from
 an event source and so is a fixed server interface implemented 
 by a client to handle these types of events.
 This function (\code{generateEventInterface})
 creates the R code for a constructor function that
 can be  used to create an instance of the event handler,
 along with knowledge of the methods, their identifiers,
 their parameter types and return types, etc.
 This is used when processing the event interface elements of a
 type library in  \code{\link{generateInterface}}.

}
\usage{
generateEventInterface(info, id, lib, defaultClassName = "CompiledCOMServer",
                        verbose = FALSE)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{info}{the \code{\link{ITypeInfo-class}} object
   (typically of class \code{\link{ITypeInfoDispatch-class}})
    that describes the event interface. It is from this that 
   we get its methods.}
  \item{id}{a string (character vector of length 1) giving 
     the human-readable name of the interface, e.g
   \code{"WorkbookEvents"}}
  \item{lib}{the \code{\link{ITypeLib-class}} object
   used to resolve references to other DCOM interfaces/types}
  \item{defaultClassName}{the name of the class to use
    as the base class when defining new R classes.}
  \item{verbose}{a logical value, with \code{TRUE} indicating that
    output reporting progress and the current activity should be displayed
    on the console.}
}
\value{
An object of class \code{\link{RCOMEventInterfaceDefinition-class}}.
}


\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
 \code{\link{generateInterface}}
 \code{\link{generateClass}}
}
%\examples{}
\keyword{programming}
\keyword{interface}
\concept{meta-programming}
\concept{distributed computing}


