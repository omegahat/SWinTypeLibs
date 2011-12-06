\name{ITypeLib-class}
\docType{class}
\alias{ITypeLib-class}
\alias{IContainingTypeLib-class}
\alias{names,ITypeLib-method}
\alias{length,ITypeLib-method}
\alias{[[,ITypeLib-method}
\alias{[[,ITypeLib,numeric-method}
\alias{[[,ITypeLib,character-method}
\alias{[,ITypeLib,numeric-method}
\alias{[,ITypeLib,logical-method}
\alias{[,ITypeLib,character-method}
\alias{$,ITypeLib,character-method}
\alias{sapply,ITypeLib-method}
\alias{lapply,ITypeLib-method}
\title{Class "ITypeLib"}
\description{
 Handle to an internal C-level type library (ITypeLib object
 in the Microsoft API) which contains meta-information about
 classes, methods, enumerations, interfaces, etc.

 The \code{IContainingTypeLib} class is a special type of
 \code{ITypeLib}. For the most part, it behaves much the same
 as the \code{ITypeLib}, but it is slightly different.
 It is obtained via a different source than loading the type library
 from a file. Rather it is computed from an \code{ITypeInfo} object
 which may come directly from a DCOM instance.  One can fetch the
 more complete \code{ITypeLib} via a call to \code{\link{LoadTypeLib}}.
 This \code{IContainingTypeLib} is defined in order to identify to the
 user that the specific type library is not the same as a regular 
 \code{ITypeLib}.
}
\section{Objects from the Class}{
 These are created using \code{\link{LoadTypeLib}}.
}
\section{Slots}{
  \describe{
    \item{\code{ref}:}{Object of class \code{"externalptr"}.
     The address of the C-level object.}
    \item{\code{fileName}:}{character string (i.e. of length 1).
      This gives the name of the file containing the type library 
      information.}
  }
}

\section{Methods}{
  \describe{
    \item{length}{\code{signature(x = "ITypeLib")}: compute the number of elements
      in the type library.}
  }
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
 \code{\link{ITypeInfo-class}}
}

\examples{
}
\keyword{classes}
