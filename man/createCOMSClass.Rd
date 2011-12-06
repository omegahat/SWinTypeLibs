\name{createCOMSClass}
\alias{createCOMSClass}
\title{Generate S class mirroring COM class}
\description{

This generates a simple S class that provides structured access to a
corresponding COM class.  This function is used to ``compile'' COM
classes for S.  Using the type information for the COM class, we
generate methods for the \code{$} and \code{$<-} operators which
provide controlled and efficient access to COM objects of this class.
These methods have information about the properties and methods of the
COM class and can detect errors on the client side of the interface
before making calls to the COM server.  Additionally, the operators
use the name-identifier mappings from the type information to avoid an
extra round-trip for each \code{\link[RDCOMClient]{.COM}} call.  Also, the
operators know which elements of the COM class are properties and
which are methods and so one can use the \code{$} operator to access
both types of elements rather than using \code{[[} for properties and
\code{$} for methods.

This defines a new S4 class that is a simple extension of the
\code{\link[RDCOMClient]{COMIDispatch-class}} class in the
\code{RDCOMClient} package.  The new class is used primarily to
support methods that know the specific details of the COM class.
Objects of this new class can be created manually or implicitly in the
\code{RDCOMClient} conversion mechanism that creates COM objects from
C++ IDispatch objects.  In the latter case, the
\code{\link[RDCOMClient]{createCOMReference}} function must be able to
map the UUID of the COM object to an S4 class name that extends
\code{\link[RDCOMClient]{COMIDispatch-class}}. 
This \code{createCOMSClass} function writes this information into 
the global character vector \code{.COMSClassMap}.
}
\usage{
createCOMSClass(libEntry, className, where = NULL)
}
\arguments{
  \item{libEntry}{the \code{\link{ITypeInfo-class}} object describing the COM class.}
  \item{className}{the name to use for the new S class that is created}
  \item{where}{unused at present.}
}
\value{
The name of the new S4 class.
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}
\seealso{
 \code{\link{computeFunctionInformation}}
 \code{\link{generateOperators}}
 \code{\link{getTypeInfo}}
 \code{\link{LoadTypeLib}}
 \code{\link[RDCOMClient]{createCOMReference}}
}
\examples{
\dontrun{
lib = LoadTypeLib("C:\\\\Program Files\\\\Microsoft Office\\\\Office\\\\EXCEL9.OLB")
createCOMSClass(lib[["_Application"]], "ExcelApplication")

 library(RDCOMClient)
 e = COMCreate("Excel.Application")
 e = as(e, "ExcelApplication")
 e$Visible
 e$Visible = TRUE
 e$Workbooks$Add()
 e$CheckSpelling("Is this okey?")
}
}
\keyword{programming}

