\name{writeCode}
\alias{writeCode}
\alias{writeCode,RCOMTypeLibraryInterfaceDefinition,ANY-method}
\alias{writeCode,RCOMTypeLibraryInterfaceDefinition,character-method}
\alias{writeCode,RCOMTypeInterfaceDefinition,character-method}
\alias{writeCode,RCOMTypeInterfaceDefinition,ANY-method}
\alias{writeNamespace}

\title{Write the generated code to a stream/connection.}
\description{
 These methods handle writing the generated code to 
 a directory, file or connection so that the code 
 can be used in future R sessions.
}
\usage{
writeCode(def, file = stdout(), context = NULL,
           prepend = character(),
           enumsAsValues = TRUE, 
           namespace = "NAMESPACE",
           ...)
writeNamespace(ns, nsFile = "NAMESPACE", copy = character())
}
\arguments{
  \item{def}{the code definitions to be written.}
  \item{file}{the connection to which the code will be written.
    If this is a directory, individual files are created for each class.}
  \item{context}{contextual information that can govern the output of the code.}
  \item{prepend}{a character vector giving the fully-qualified names of files whose contents are
    to be inserted before the generated code. If the output is to a directory, these
    files are just copied directly there as separate files with the same names as given.}
  \item{enumsAsValues}{a logical value that is passed as the third argument to 
    \code{writeEnums} and this controls whether the individual name-value
    pairs within an enumeration definition are defined as separate, individual R variables.}
 \item{namespace}{a character vector giving the name of the file to which the 
     name space information, e.g export, exportClasses, etc., should be written.
    This defaults to the usual NAMESPACE. 
    To avoid having the name space information written, specify either \code{NA}
    or \code{character()}.
  }
  \item{\dots}{any additional information that is passed onto the different methods.}

  \item{ns}{the name space information returned from \code{writeCode} which is to be
         written to the file given by \code{nsFile}.}
  \item{nsFile}{the name of the file to which the name space information is to be written.
    This is a character vector.}
  \item{copy}{the fully-qualified names of any addtional files whose
    contents are to be copied into the resulting NAMESPACE file.
    These are appended to the end of the newly generated
    NAMESPACE file.
  }
}
\details{
}
\value{

}
\references{
 \url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}
\note{}


\seealso{
 \code{\link{define}}
 \code{\link{computeFunctionInformation}}
}
\examples{

}
\keyword{interface}
\keyword{programming}
\concept{reflection}

