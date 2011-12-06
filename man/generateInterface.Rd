\name{generateInterface}
\alias{generateInterface}
\title{Generate S code to interface to classes in a DCOM library.}
\description{
 This is the top-level entry point for generating S code 
(classes and accessor methods) for interfacing to 
 DCOM classes described in a type library.
 Perhaps it should be called \code{generateTypeLibInterface}.

 It processes all of the "dispatch" interfaces described
 in the library (or just those specified in the \code{classes}
 argument) and generates an S4 object describing the interface
 to that class and its properties and methods.
}
\usage{
generateInterface(lib, classes = names(lib), events = TRUE, 
                  enums = TRUE, coclass = TRUE,
                  defaultClass = "CompiledCOMIDispatch",
                  defaultListClass = c("COMList", "COMTypedList"),
                  class = "RCOMTypeLibraryInterfaceDefinition",
                  verbose = FALSE, processExternalClasses = TRUE,
                  computeReferencedClasses = TRUE)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{lib}{the ITypeLib object containing all the DCOM class descriptions.
   This is typically obtained via a call to \code{\link{LoadTypeLib}}, 
    either loading a type library file directly or from a COMIDispatch
    object which is an existing instance of a DCOM object.
   }
  \item{classes}{a character or integer vector giving the identities of the 
   subset of elements in the type library that are to be processed.
   This allows callers to generate interfaces to a smaller set of the DCOM classes,
   as some may not be relevant.}
 \item{coclass}{either a logical value, a character vector or a list.
   If this is a logical, the value \code{TRUE} indicates to process all
   the co-classes in the library, and if \code{FALSE} to ignore them.
   A character vector names the co-class elements in the library that
   should be processed. And a list provides those elements explicitly.
  }
  \item{events}{a logical value indicating whether to generate interfaces to the
   event interfaces in the type library.}
  \item{enums}{a logical value indicating whether to generate interfaces to the
   enumerations defined in the type library.}
  \item{defaultClass}{a character vector of length 1 (i.e. a string) giving the name 
    of the S4 class which is to act as the base class for all of the newly generated
    classes.   (This does not apply to DCOM classes that are lists. See \code{defaultListClass}.)}
  \item{defaultListClass}{a string giving the name of the base class to use for
    DCOM classes that are identified as lists. This is typically COMList, but
    can be specified differently in order to, for example, handle lists that use zero-based
    counting (e.g. the links in the Internet Explorer type library).}
  \item{class}{a string giving the name of the class to create to represent the
   information. This is the class of the return value from this function.}
  \item{verbose}{a logical value, with \code{TRUE} indicating that
    output reporting progress and the current activity should be displayed
    on the console.}
  \item{processExternalClasses}{a logical value indicating whether to 
   generate interfaces to any DCOM types that are referenced as 
   types for parameters or return values that are not actually
   defined within the type library being processed, but rather 
   another library on which this one depends.  If this is \code{TRUE},
   \code{\link{RCOMTypeLibraryInterfaceDefinition-class}} 
   objects are defined for each of the referenced libraries and interface 
   descriptions generated  for each of the  referenced types within
   them.}
 \item{computeReferencedClasses}{a logical indicating whether to
   find all the types that are referenced in parameters,
   return types, properties, etc. and include those in the result.
   These are needed for a complete interface to the specified
   classes, methods and properties.
 }
}

\value{
 An object of class given by the \code{class} argument.
 This is instantiated with a single field which are the
 descriptions of each of the elements.
}
\references{
 \url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{writeCode}}
}
\examples{
\dontrun{
 lib = LoadTypeLib("C:/Microsoft Office/OFFICE11/Excel.EXE")
 def = generateInterface(lib)
 writeCode(def, "Excel.S")
}
}
\keyword{programming}
\keyword{interface}

