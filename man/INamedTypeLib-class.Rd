\name{INamedTypeLib-class}
\docType{class}
\alias{INamedTypeLib-class}
\alias{[,INamedTypeLib,numeric-method}
\alias{[,INamedTypeLib,character-method}
\alias{[[,INamedTypeLib,numeric-method}
\alias{[[,INamedTypeLib,character-method}
\alias{getElements,INamedTypeLib-method}
\alias{getTypeLibUUIDs,INamedTypeLib-method}
\alias{names,INamedTypeLib-method}
\alias{show,INamedTypeLib-method}

\title{Class "INamedTypeLib" ~~~ }
\description{
 This class is merely an "efficient" version of the
 \code{\link{ITypeLib-class}} that is used when making
 intensive use of the contents of the library.
 The idea is simply that we cache the commonly accessed
 elements of the library: the names of the elements,
  their UUIDs, the top-level \code{\link{ITypeInfo-class}} elements
  in the library. Then, we provide methods that replace
  the regular access to this information.
  In this way, we avoid recomputing the information
  repeatedly. When processing an entire type library
  as in a call to \code{\link{generateInterface}}, this 
  can be very significant.
}
\section{Objects from the Class}{
 Objects should be created using the constructor
 \code{link{INamedTypeLib}} which computes the information
 from the regular \code{\link{ITypeLib-class}} object and 
 caches it in the new object.
}
\section{Slots}{
  \describe{
    \item{\code{elNames}:}{Object of class \code{"character"}
    these are the names of the top-level elements of the type library.
     This is the same value as calling \code{names} on a regular
      \code{\link{ITypeLib-class}} object, but instead these are cached.}
    \item{\code{infoEls}:}{Object of class \code{"list"} ~~ }
    \item{\code{uuids}:}{Object of class \code{"character"} ~~ }


    \item{\code{ref}:}{Object of class \code{"externalptr"}  inherited from
        \code{\link{ITypeLib-class}}. }
    \item{\code{fileName}:}{Object of class \code{"character"} inherited from
        \code{\link{ITypeLib-class}}. }
  }
}
\section{Extends}{
Class \code{"ITypeLib"}, directly.
}
\section{Methods}{
  \describe{
    \item{[}{\code{signature(x = "INamedTypeLib", i = "numeric")}:  }
    \item{[}{\code{signature(x = "INamedTypeLib", i = "character")}:  }
    \item{[[}{\code{signature(x = "INamedTypeLib", i = "numeric")}:  }
    \item{[[}{\code{signature(x = "INamedTypeLib", i = "character")}: ... }
    \item{getElements}{\code{signature(x = "INamedTypeLib")}: ... }
    \item{getTypeLibUUIDs}{\code{signature(lib = "INamedTypeLib")}: ... }
    \item{names}{\code{signature(x = "INamedTypeLib")}: ... }

    \item{show}{\code{signature(object = "INamedTypeLib")}: }
  }
}

\references{
 \url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
 \code{\link{ITypeLib-class}}
 \code{\link{generateInterface}}
}
\examples{
}
\keyword{classes}
