\name{getElements}
\alias{getElements}
\alias{getElements,ITypeInfo-method}
\alias{getElements,ITypeLib-method}
\alias{getElements,ITypeInfoAlias-method}
\alias{getElements,ITypeInfoCoClass-method}
\alias{getEnum}
\alias{getEnum,ITypeInfoEnum-method}
\title{General mechanism for retrieving elements of a Type Library element}
\description{
  \code{getElements} is a function that serves as a
  general convenient
  way to extract the "relevant" elements from
  a type library or an element within a type library.
  For a type library, the relevant elements are  
  descriptions of all the classes in the form
  of a list of \code{\link{ITypeInfo-class}}.
  When applied to one of these
  \code{\link{ITypeInfo-class}} itself,
  \code{getElements} accesses the sub-elements
  such as the methods and property accessors
  for a \code{\link{ITypeInfoDispatch-class}}
  object, or the collection of
  \code{\link{ITypeInfoDispatch-class}}
  for a \code{\link{ITypeInfoCoClass-class}} object.
  If the object is an enumeration (\code{\link{ITypeInfoEnum-class}}),
  this function returns the collection of associated named constants.


  \code{getEnum} is the actual function
  that processes \code{ITypeInfoEnum-class} elements within a
  type library to get the "elements".
  \code{\link{getFuncs}} is used for 
  \code{ITypeInfoDispatch-class} library elements.
}
\usage{
getElements(x, recursive = FALSE)
getEnum(x)
}
\arguments{
  \item{x}{ an object of class \code{ITypeInfo-class}}
  \item{recursive}{a logical value }
}
\value{
 \code{getElements} returns a list.
 Depending on the type of the \code{\link{ITypeInfo-class}}, 
 different information is  returned.

 \code{getEnum} returns a named integer vector.
 The names of the vector's elements
 give the symbolic names of the constants as used in the C/C++ code.
 These are more readable and more robust than using separate values.
 The values are the values of the symbolic constants
 used in the C/C++ code.
 See \code{\link[RDCOMClient]{EnumValue}}.
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{getFuncs}}
 \code{\link{LoadTypeLib}}
}
\examples{
\dontrun{
  e = getCOMInstance("Excel.Application")
  lib = LoadTypeLib(e)

  # CoClass
  getElements(lib[["Application"]])

  # Methods of an IDispatch type
  getElements(lib[["_Application"]])
  #  Same thing.
  getFuncs(lib[["_Application"]])

  # Enumeration.
  getElements(lib[["Constants"]])
  getEnum(lib[["Constants"]])
}
}
\keyword{programming}
\concept{meta-programming}


