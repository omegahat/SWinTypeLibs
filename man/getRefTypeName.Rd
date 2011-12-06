\name{getRefTypeName}
\alias{getRefTypeName}
\title{Get the human-readable name of the interface identified by a reference type}
\description{
  This function processes an object of class \code{\link{TypeDescriptionRef-class}}
  and finds the name of the human-readable type by looking up the UUID 
  identified by the reference type and comparing it to those in the 
  specified type library.
  
}
\usage{
getRefTypeName(info, ref, lib, follow = !is.null(lib))
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{info}{the \code{\link{ITypeInfo-class}} object associated with
     \code{TypeDescriptionRef}. For instance, if the 
    \code{TypeDescriptionRef} is obtained as the return type or parameter type in a
     method, the \code{info} is the \code{\link{ITypeInfo-class}} object
     from which the methods were retrieved.
    }
  \item{ref}{an object of class \code{\link{TypeDescriptionRef-class}}
    or its value (i.e. the \code{reftype} field)}
  \item{lib}{the type library, of class \code{\link{ITypeLib-class}}}
  \item{follow}{a logical value indicating whether to resolve the
  type information in its native library if it is not actually
   defined within \code{lib}, but just referenced there.
   This is passed to \code{\link{getRefTypeInfo}}}
}
\value{
 A string (character vector of length 1) giving the name
  of the element in the type library identified by the
  reference type.
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
 \code{\link{getRefTypeInfo}}
 \code{\link{getTypeLibUUIDs}}
}
\examples{

}
\keyword{interface}

