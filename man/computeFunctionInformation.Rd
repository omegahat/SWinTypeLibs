\name{computeFunctionInformation}
\alias{computeFunctionInformation}
\title{Create information about a COM method}
\description{
 This function is used by \code{\link{createCOMSClass}}
when creating methods for \code{$} and \code{$<-}
for a COM class. This computes information
about the parameters for the method,
which are required and a collection of converter functions
for marshalling the S values in the call to the target
COM types.
}
\usage{
computeFunctionInformation(desc)
}
\arguments{
  \item{desc}{the \code{\link{FunctionDescription-class}} object
   describing the COM method.}
}
\value{
A list
  \item{paramNames}{the names of the different parameters}
  \item{required}{a vector giving the names of the required arguments}
  \item{converters}{not currently meaningful}
}

\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
 \code{\link{createCOMSClass}}
 \code{\link{generateOperators}}
 the \code{RDCOMClient} package
}
%\examples{}

\keyword{programming}
