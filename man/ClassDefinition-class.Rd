\name{ClassDefinition-class}
\docType{class}
\alias{ClassDefinition-class}
\alias{InlineClassDefinition-class}
\alias{NamedClassDefinition-class}
\alias{classDefinitionString}

\title{Classes representing code for defining DCOM classes.}
\description{
 These classes are used to represent pending or would-be
 R class definitions for DCOM interfaces.
 These are converted into actual commands to define R classes
 when the code is written via \code{\link{writeCode}} 
 to a file for use in future R  sessions.
 So these are not class definitions directly, but
 representation of information about how to define
 particular classes.

 \code{ClassDefinition} is a virtual class.
 This provides a slot inherited by the other two classes
 giving the name of the new class to be defined.

 \code{NamedClassDefinition} is a simple class 
 that contains information about the names of the
 class(es) that the new class extends.

 \code{InlinedClassDefinition} uses a slightly different
 representation, which is the R command (as a string)
 that defines the class.  This allows it to 
 include more complex structure in the \code{setClass}
 command such as a prototype.


 The function \code{classDefinitionString} is 
 used to convert an instance these classes
 into a string (a character vector of length 1)
 that can be used within code to define the class.
}
\section{Objects from the Class}{
 Objects can be created by calls of the form 
 \code{new("InlineClassDefinition", ...)}
 and \code{new("NamedClassDefinition")}.
} 
\section{Slots}{
  \describe{
    \item{\code{code}:}{Object of class \code{"character"}.
     For \code{NamedClassDefinition},
      this is a character vector giving
        the name of the new class being defined,
         and the class(es) that it contains.
       This is conveted into a definition as
        \code{setClass(x@className, contains = x@code[2])},
       so currently only the second element is really used.

       For \code{InlineClassDefinition}, again the \code{className}
       gives the name of the new class being defined.
       The code slot is a string that gives the 
       R command to define the new class. As such, it is entirely
       free form but must evaluate correctly when parsed and evaluated
       by the R interpreter.
     }
    \item{\code{className}:}{Object of class \code{"character"}.
         This is a string (i.e. vector of length 1) that gives the 
         name of the new class being defined.}
  }
}
\section{Methods}{
No methods defined with class "ClassDefinition" in the signature.
}
\references{
 \url{http://msdn.microsoft.com/library}
 \url{http://www.omegahat.org/RWinTypeLibs}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
 \code{\link{classDefinitionString}}
 \code{\link{generateInterface}}

 \code{getListClassName}
 \code{computeClassName}
}
\examples{
 new("NamedClassDefinition", 
      className = "Workbook",
      code = "CompiledCOMIDispatch")

 cmd = "setClass('Workbooks', contains = 'COMNamedTypedList',
                 prototype = list(name = '_Workbook'))"
 new("InlineClassDefinition", 
      className = "Workbooks",
      code = cmd)
}
\keyword{classes}
