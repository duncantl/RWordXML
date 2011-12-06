\name{RWordXMLDoc-class}
\Rdversion{1.1}
\docType{class}
\alias{RWordXMLDoc-class}

\title{Representation of main XML document within a .docx file}
\description{This class represents a parsed XML document that originates
  from reading the primary document within a Word .docx file.
  This is a regular \code{XMLInternalDocument} from the \code{XML}
  package, but has the additional class information so that we can
  operate on it via methods that understand the nature of the content,
  i.e. Word structure.
 }
\section{Objects from the Class}{A virtual Class: No objects may be created from it.}
\section{Slots}{
  \describe{
    \item{\code{.S3Class}:}{Object of class \code{"character"} this an
      opaque data type inherited from the \code{XML} package.}
  }
}
\section{Extends}{
Class \code{"\linkS4class{XMLInternalDocument}"}, directly.
Class \code{"\linkS4class{XMLAbstractDocument}"}, by class "XMLInternalDocument", distance 2.
Class \code{"\linkS4class{oldClass}"}, by class "XMLInternalDocument", distance 3.
}
\section{Methods}{
  \describe{
    \item{coerce}{\code{signature(from = "RWordXMLDoc", to = "WordArchive")}: ... }
	 }
}
\references{ ~put references to the literature/web site here ~ }
\author{Duncan Temple Lang}

\seealso{
 \code{\link{WordArchive-class}}
}
\examples{
f = system.file("SampleDocs", "sampleLists.docx", package = "RWordXML")
doc = wordDoc(f)
x = doc[[getDocument(doc)]]
y = doc[["word/document.xml"]]  # note that this doesn't have the special class.
class(x)
class(y)
class(doc[["/word/document.xml"]])  # this one does because of the /.
                                    #  We may get around to changing this!
}
\keyword{classes}
