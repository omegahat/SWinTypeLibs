\name{queryTypeLibPath}
\alias{queryTypeLibPath}
\alias{queryTypeLibPath,COMIDispatch,ANY-method}
\alias{queryTypeLibPath,ITypeInfo,ANY-method}
\alias{queryTypeLibPath,ITypeLib,ANY-method}
\alias{queryTypeLibPath,ITypeLibIdentifier,ANY-method}
\alias{queryTypeLibPath,IContainingTypeLib,ANY-method}
\alias{queryTypeLibPath,character,integer-method}
\alias{queryTypeLibPath,ITypeLibIdentifier-method}
\title{Find the full path of the associated type library.}
\description{
 This function (and its methods) can locate the full 
 name of the file that is the type library associated
 with an appropriate object  in R.
 It can be used on DCOM objects (of class \code{COMIDispatch}),
 \code{ITypeInfo} objects derived from a DCOM object or from
 a type library, or from a type library itself.
 Additionally, it can also be called with the UUID (specifically
 the GUID) identifying a class/interface and it consults the 
 Windows' registry to find the associated type library.
 It is a convenient way to find the actual file itself
 from an instance of a DCOM object,
 so that it can be loaded directly in other sessions
 without having to create the DCOM object first.
}
\usage{
queryTypeLibPath(guid, version, lcid = 0)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{guid}{the object identifying the type library.
   This can be a \code{\link[RDCOMClient]{COMIDispatch}},
   or \code{link{ITypeInfo}}, or a UUID or character string giving the 
   UUID.}
  \item{version}{if specified, an integer vector of length 2 giving the 
    major and minor version numbers (in that order) of the library of interest.
  }
  \item{lcid}{a numeric value giving the locale identifier to identify the
   particular version of the library of interest.  The default works well.}
}
\value{
 A character vector or an error.
 The character vector is the full name of the file identifying
 the type library for the given object.
}
\references{
 \url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
  \code{\link{LoadTypeLib}}
}
\examples{
\dontrun{
  library(RDCOMClient)
  e = COMCreate("Excel.Application")
  fileName = queryTypeLibPath(e)

  lib1 = LoadTypeLib(fileName)
  lib2 = LoadTypeLib(e)
  # lib1 and lib2 should be the same.
}
}
\keyword{programming}
\keyword{interface}
