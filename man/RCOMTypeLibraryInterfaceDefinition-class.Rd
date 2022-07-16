\name{RCOMTypeLibraryInterfaceDefinition-class}
\docType{class}
\alias{RCOMTypeLibraryInterfaceDefinition-class}

\title{Class to represent R interface to a Type Library}
\description{
 This class is used to represent all the information 
 gathered in processing the contents of a DCOM-related
 type library
 to define an R interface to all the classes, enumeration
 elements, event handlers, etc. in that type library.
 This is created by a call to \code{\link{generateInterface}}
 and the result is an object that contains information
 about the necessary class definitions, property and method
 accessor functions, event handler constructors,
 and enumeration definitions.
 These can then be written to a connection for use
 in future R sessions via \code{\link{writeCode}}.  
 Some incomplete code allows it to be used
 directly in the current session, or "just in time",
 via the \code{define} function.
}
\section{Objects from the Class}{
 Objects are typically created using  the function
 \code{\link{generateInterface}}.
}
\section{Slots}{
  \describe{
    \item{\code{types}:}{Object of class \code{"list"}
          a collection of objects describing an R-level interface
       to the corresponding DCOM type in the type library.
       These elements are of class \code{\link{RCOMDispatchInterfaceDefinition-class}}.}
    \item{\code{enums}:}{Object of class \code{"list"}
       a collection of objects describing an enumeration definition in the type library.
       Each object is merely the named integer vector giving the symbolic names and the  
       corresponding value. }
    \item{\code{events}:}{Object of class \code{"list"}
       a collection of objects describing the interface to an event source
       for responding to  its event announcements or triggers.
       Each element is an object of class \code{\link{RCOMEventInterfaceDefinition-class}}.}
    \item{\code{externalClasses}:}{Object of class \code{"list"}.
        This is used to store a collection of \code{RCOMTypeLibraryInterfaceDefinition}
        objects.  There is one for each of the libraries and their types that are referenced
        within the original top-level library. This happens when we process
        a type library to create an R interface to it
        (via \code{\link{generateInterface}}) and encounter types for parameters
        and return values of methods and properties that are not actually
        defined in the original type library, but rather in secondary libraries.
        We can generate interfaces to these types in those libraries in the same
        way we do for the original top-level/initial library and we store these
        interfaces in this slot.}
      \item{\code{coclass}:}{a list.  The co-class interfaces of
	interest in this library.}
      \item{\code{referencedClasses}:}{a list.  This identifies types in
      the library that are return types, parameter types, or property
      types that were encountered indirectly when processing the
      specified elements of the library.}
      \item{\code{typeEventMap}:}{a list.  Information about the events
	in the classes of this interface}
      \item{\code{uuids}:}{}      
  }
}
\section{Methods}{
  \describe{
    \item{writeCode}{\code{signature(def = "RCOMTypeLibraryInterfaceDefinition", file = "character")}: write the R code contained in this interface definition 
   to a directory or file.  This is responsible for creating the necessary
   connections to the files and then writing the code.}
    \item{writeCode}{\code{signature(def = "RCOMTypeLibraryInterfaceDefinition", file = "ANY")}: write this code to the specified connection.}
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

 It may be useful to look at non-exported functions
 in \code{generate.S} and \code{referencedClasses.S}
 to see how the entire thing works and how it can be
 changed.  This interface is not cast in stone. 
 We want others to be able to explore alternatives,
 ideally by extending or parameterizing the mechanism
 rather than presenting an entirely parallel but similar
 mechanism.
 
}
%\examples{}
\keyword{classes}
