\name{getTypeInfo}
\alias{getTypeInfo}
\alias{getTypeInfo,ITypeLib,numeric-method}
\alias{getTypeInfo,ITypeLib,missing-method}
\alias{getTypeInfo,ITypeLib,character-method}
\alias{getTypeInfo,ITypeInfo,TypeDescriptionRef-method}
\alias{getTypeInfo,COMIDispatch-method}
\title{Reference to ITypeInfo Object}
\description{
 This is used to get the C-level representation
of an \code{ITypeInfo} object. Currently, one gets this from the \code{ITypeLib}
or from a \code{\link[RDCOMClient]{COMIDispatch}} object.
From this object, one can query the meta-information it contains
such as method definitions, etc.
}
\usage{
getTypeInfo(obj, which)
}
\arguments{
  \item{obj}{the object from which to get the \code{ITypeInfo} object.}
  \item{which}{the index (1-based) of the element within the \code{ITypeLib}
   whose \code{ITypeInfo} object is desired.
   If this is omitted, all the elements in the type library are returned.}
}
\details{

}
\value{
 An object of class \code{ITypeInfo}
 which contains a slot giving the address of 
 the C-level object representing the \code{ITypeInfo} object.
}

\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
 \code{\link{LoadTypeLib}}
 \code{\link{LoadTypeLib}}
}

\examples{
if(file.exists("c:\\\\WINNT\\\\system32\\\\wmpui.dll")) {
  l = LoadTypeLib("c:\\\\WINNT\\\\system32\\\\wmpui.dll")
  
  print(getTypeInfo(l, 3))
}  
}
\keyword{programming}
%\keyword{reflectance}
