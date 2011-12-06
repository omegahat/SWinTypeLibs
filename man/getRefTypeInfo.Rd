\name{getRefTypeInfo}
\alias{getRefTypeInfo}
\alias{getRefTypeInfo,ITypeInfo,numeric-method}
\alias{getRefTypeInfo,ITypeInfo,TypeDescriptionRef-method}

\title{Find the ITypeInfo object for the type given in a TypeDescriptionRef object.}
\description{
 This function resolves the  \code{\link{ITypeInfo-class}} element in a type library
 for a given \code{\link{TypeDescriptionRef-class}} object.
 These \code{\link{TypeDescriptionRef-class}}  objects arise
 in the description of parameter types and return types
 of DCOM methods  via \code{\link{getFuncs}} and we 
 need to resolve them to find their actual definitions 
 in terms of DCOM interfaces.
 This function performs that action of mapping the
 information in the \code{\link{TypeDescriptionRef-class}} 
 object  to a \code{\link{ITypeInfo-class}}.
}
\usage{
getRefTypeInfo(info, href, lib = NULL, follow = !is.null(lib))
}
\arguments{
  \item{info}{the \code{\link{ITypeInfo-class}} object in which to
   resolve the reference.}
  \item{href}{the identifier for the reference, typically
   the \code{reftype} slot of a \code{\link{TypeDescriptionRef-class}}
   object}
  \item{follow}{a logical value indicating whether to resolve the
    referenced type across type libraries. This arises when 
    library A refers to this type as a parameter or return value
    but it is actually defined in a type library, B say.  We can resolve
    this in library A, but the results may be generic (e.g. 
     an enumeration may have a name Constants rather than the more specific
     name).  If we resolve the type in library B, then we can get the more 
    specific \code{\link{ITypeInfo-class}} object.}
  \item{lib}{an object of class \code{\link{ITypeLib-class}}.
    This is connected to the \code{follow} argument the purpose of which you should
    understand before reading this.  This argument provides the library
    instance in which we encountered the reference (given by \code{href}).
    If it is provided and \code{follow} is \code{TRUE}, then we will 
    use this \code{lib} value to determine if we the
    \code{ITypeInfo} object is defined in a different library or this one.
    If it is defined in a different library and \code{follow} is \code{TRUE}, 
    then we will resolve it there. So, this \code{lib} parameter allows us to 
    compare the origin of the definition of the type and resolve it appropriately.
   }
}
\value{
 An object of some class derived  from \code{\link{ITypeInfo-class}}
 depending on the type the reference identifies.
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link{getTypeInfo}}, specifically the method for
  \code{ITypeInfo, TypeDescriptionRef}.
 \code{\link{getRefTypeName}}
}
\examples{
\dontrun{

}
}
\keyword{programming}
\keyword{interface}
\concept{meta-programming}