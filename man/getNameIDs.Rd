\name{getNameIDs}
\alias{getNameIDs}
\alias{getNameIDs,FunctionDescription,ITypeInfo-method}
\alias{getNameIDs,ITypeInfo,ANY-method}
\alias{getNameIDs,character,ITypeInfo-method}
\alias{getNameIDs,character,COMIDispatch-method}
\alias{getNameIDs,numeric,ITypeInfo-method}
\title{Compute the numeric identifiers for DCOM (method) names.}
\description{
 This function uses the type library to compute the 
 mapping of names, such as method names, to the 
 numeric identifiers used in dispatching a
 a call using those names. 

 The \code{.COM} functionality used to invoke
 a method first sends a query to resolve the
 identifiers for the names in the method call.
 Then it uses these to actually make the call.
 Therefore, precomputing these identifiers  
 and using them in calls can improve the 
 efficiency of the call by avoiding the first step.

 This information is also used when we implement a DCOM server
 in R. When other clients query the identifier for a name of a method
 in our server's interface, we must respond appropriately.
 When implementing already defined, established interfaces,
 again precomputing the name-identifier map helps.

 This is a generic function that operates on various different
 pairs of input types.

}
\usage{
getNameIDs(obj, context)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{obj}{the object which is to be interrogated for the
    name-identifier mapping.}
  \item{context}{support information in which some methods will resolve the 
    name-identifier connections.
    For example,  }
}
\value{

}

\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link[RDCOMClient]{.COM}} and the \code{ids} argument.
 \code{\link[RDCOMServer]{COMInterfaceServerInfo-class}} and
        its \code{ids} slot.
}
\examples{
\dontrun{

     # FunctionInvokeDescription and an ITypeInfo.
   funcs = getFuncs(lib[["_Application"]])
   getNameIDs(funcs$Quit, lib[["_Application"]])

    # include parameters, although these will come back as 0, 1, ...
   getNameIDs(funcs$GetSaveAsFilename, lib[["_Application"]])

    # Instead of getting the FunctionInvokeDescription object first,
    # let getNameIDs do the work.  Works for arbitrary objects
    # in the ITypeInfo, not just FunctionInvokeDescription objects,
    # but property set and get methods, etc.
   getNameIDs("Quit", lib[["_Application"]])

    # Get the name mappings for all the elements of an ITypeInfo,
    # e.g. all functions all at once.
   getNameIDs(lib[["_Application"]])   
}
}
\keyword{programming}
\keyword{interface}
\concept{meta-programming}

