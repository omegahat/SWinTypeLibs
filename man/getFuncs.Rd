\name{getFuncs}
\alias{getFuncs}
\alias{getFuncs,ITypeInfo-method}
\alias{getFuncs,COMIDispatch-method}
\title{Meta-Information about Functions}
\description{
 This is a generic function that retrieves information describing 
 the collection of different functions or routines
 within a type library or an (D)COM object.
}
\usage{
getFuncs(x, lib = NULL)
}
\arguments{
  \item{x}{the object in which to lookup the 
   functions.
   This can be a \code{\link{ITypeInfo-class}} or a 
   \code{\link[RDCOMClient]{COMIDispatch-class}}.
   The former is computed from the \code{ITypeLib} object.}
  \item{lib}{an \code{\link{ITypeLib-class}} object 
   associated with the \code{\link{ITypeInfo-class}} (i.e. the value of 
   \code{x})
   object in which the functions are to be found.  It is often helpful
   to be able to specify this to avoid extra computations.
   Specifically, we can use cached values for a library
   via the \code{\link{INamedTypeLib-class}} which 
   pre-fetches the primary components of the library and avoids
   the overhead of recomputing them each time they are needed.
% and also to provide a referenced library.
  }
}
\value{
 A list of \code{\link{FunctionDescription-class}}
 objects.
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}
\seealso{
 \code{\link{getElements}}
 \code{\link{LoadTypeLib}}
 \code{\link[RDCOMClient]{COMCreate}}
 \code{\link[RDCOMClient]{.COM}}
}

\examples{

}
\keyword{programming}
%\keyword{reflectance}
