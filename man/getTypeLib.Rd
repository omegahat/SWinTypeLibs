\name{getTypeLib}
\alias{getTypeLib}
\alias{getTypeLib,character-method}
\alias{getTypeLib,COMIDispatch-method}
\alias{getTypeLib,ITypeInfo-method}
\title{Get a handle to the type library associated with an object}
\description{
 This  function and its methods are very similar to \code{\link{LoadTypeLib}}
 in that they return a \code{\link{ITypeLib-class}} object
 containing a reference to a C-level  type library.
 Currently, the only difference is that the method for 
 a COMIDispatch object or a ITypeInfo object get the containing
 type library associated with the corresponding C-level object.
 And this is an object of class \code{\link{IContainingTypeLib-class}}
 as this is not necessarily the same as a library that is explicitly
 loaded from a file. As a result, \code{\link{LoadTypeLib}} is
 perhaps the preferred function.
}
\usage{
getTypeLib(x)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{x}{the object from which to get the type library information.
   This is typically an object of class \code{\link{ITypeInfo-class}}
   or a \code{\link[RDCOMClient]{COMIDispatch-class}} object.
  }
}
\value{
 An object of class \code{{ITypeLib-class}}
 or \code{\link{IContainingTypeLib-class}}.
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{LoadTypeLib}}
 \code{\link{LoadRegTypeLib}}
}
%\examples{}
\keyword{interface}
\keyword{programming}

