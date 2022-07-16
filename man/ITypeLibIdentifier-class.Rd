\name{ITypeLibIdentifier-class}
\docType{class}
\alias{ITypeLibIdentifier-class}

\title{Information identifying a Type Library}
\description{
 This class allows us to compute information about a loaded type library
(see \code{\link{LoadTypeLib}} and \code{\link{LoadRegTypeLib}})
 and find out information about it so that we can load it
 directly without needing an instance of any of the classes
 it describes.  This allows us to load it directly using 
 the Windows registry.
}
\section{Objects from the Class}{
Objects of this class are usually created using a call
to \code{\link{getTypeLibInfo}}.
One can create instances of this class as templates
to be passed to \code{\link{LoadTypeLib}}
with wildcard information about the version numbers, etc.
See the documentation for the C++ routine
\code{LoadRegTypeLib}.
}
\section{Slots}{
  \describe{
    \item{\code{guid}:}{Object of class \code{"character"},
        the GUID string identifying the type library uniquely. }
    \item{\code{version}:}{Object of class \code{"integer"},
        an integer vector of length 2 giving the major and minor
       version numbers.}
    \item{\code{lcid}:}{Object of class \code{"numeric"},
       a single numeric value identifying the locale used
       for this type library.}
    \item{\code{flags}:}{Object of class \code{"integer"},
     a scalar value giving information about the properties of
     the library. This is a bit-wise OR'ing of elements from
     the enumeration 
     \code{c(restricted = 1, control = 2, hidden = 4, hasdiskimage =8)}.
     If 8 is not in the value, there is no file on the hard drive 
     associated with this type library. The control bit
     if on indicates that the type library relates to visual controls.
     1 and 4 refer to whether the library is visible to regular users.
     For more information, see the MSDN documentation for LIBFLAGS
     and TLIBATTR.}
    \item{\code{syskind}:}{Object of class \code{"integer"}
     a named integer which is really an enumeration 
     identifying the type of system in which the library 
     is installed. This is one of the values in the vector
      \code{c(WIN16 = 0, WIN32 = 1, MAC = 2)}
     and should be reasonably clear.}
  }
}
\section{Methods}{
No methods defined with class "ITypeLibIdentifier" in the signature.
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
 \url{http://www.geocities.com/cool_ranju/tlibs.html}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
 \code{\link{getTypeLibInfo}}
 \code{\link{LoadTypeLib}}
 \code{\link{LoadRegTypeLib}}
}
%\examples{}
\keyword{classes}
