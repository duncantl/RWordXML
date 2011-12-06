\name{WordArchive-class}
\Rdversion{1.1}
\docType{class}
\alias{WordArchive-class}
\alias{[[,WordArchive,character,missing-method}
\alias{addToDoc,WordArchive,character-method}
\alias{addToDoc,WordArchive,XMLInternalNode-method}
\alias{coerce,character,WordArchive-method}
\alias{coerce,XMLInternalDocument,WordArchive-method}
\alias{coerce,XMLInternalElementNode,WordArchive-method}
\alias{coerce,RWordXMLDoc,WordArchive-method}
\alias{getDocument,WordArchive-method}
%\alias{getImages,WordArchive-method}
%\alias{getListNodes,WordArchive-method}
%\alias{getTableNodes,WordArchive-method}
\alias{hyperlinks,WordArchive-method}
%\alias{lists,WordArchive-method}
%\alias{tables,WordArchive-method}
\alias{toc,WordArchive-method}



\title{WordArchive representing a .docx file}
\description{This is a class that represents a .docx file and allows
  access to the individual elements within that archive.  We can access
  the different elements in a structured manner by knowing this is a
  Word file rather than a generic Office Open XML document.
  We do, however, inherit many of the facilities from this parent class \code{OOXMLArchive}.
}
\section{Objects from the Class}{
  Most instances of this class are created via a call to\code{\link{wordDoc}}.
}
\section{Slots}{
  \describe{
    \item{\code{.Data}:}{Object of class \code{"character"} ~~ }
    \item{\code{classes}:}{Object of class \code{"character"} names of
      the classes that this implements. (These are available from the S4
      hierarchy but are here also for some experimental purposes.)}
    \item{\code{elements}:}{Object of class \code{"character"} the names
    of the elements in the .docx file when it was read, i.e. when
    the \code{WordArchive} object was constructured, or updated when we
    add elements.}
  }
}
\section{Extends}{
Class \code{"\linkS4class{OOXMLArchive}"}, directly.
Class \code{"\linkS4class{ZipFileArchive}"}, by class "OOXMLArchive", distance 2.
Class \code{"\linkS4class{character}"}, by class "OOXMLArchive", distance 3.
Class \code{"\linkS4class{ZipArchive}"}, by class "OOXMLArchive", distance 3.
Class \code{"\linkS4class{vector}"}, by class "OOXMLArchive", distance 4.
Class \code{"\linkS4class{data.frameRowLabels}"}, by class "OOXMLArchive", distance 4.
}
\section{Methods}{
  \describe{
    \item{[[}{\code{signature(x = "WordArchive", i = "character", j =
	"missing")}: access an element in the .docx archive, parsing it
      if is an XML document. }
    \item{addToDoc}{\code{signature(archive = "WordArchive", node =
	"character")}: add XML content to the primary XML document within the archive.}
    \item{addToDoc}{\code{signature(archive = "WordArchive", node =
	"XMLInternalNode")}: add an already created XML node to the primary XML document within the archive. }
    \item{getDocument}{\code{signature(doc = "WordArchive")}: ... }
    \item{getImages}{\code{signature(doc = "WordArchive")}: ... }
    \item{getListNodes}{\code{signature(doc = "WordArchive")}: ... }
    \item{getTableNodes}{\code{signature(doc = "WordArchive")}: ... }
    \item{hyperlinks}{\code{signature(doc = "WordArchive")}: ... }
    \item{lists}{\code{signature(x = "WordArchive")}: get an R list of the
      content lists within the primary document.  }
    \item{tables}{\code{signature(x = "WordArchive")}: get a list of the
      tables within the primary document in the Word archive.}
    \item{toc}{\code{signature(file = "WordArchive")}: ... }

    \item{coerce}{\code{signature(from = "character", to =
    "WordArchive")}:  coerce a file name to a \code{WordArchive}, same as \code{wordDoc}.}
    \item{coerce}{\code{signature(from = "XMLInternalDocument", to =
    "WordArchive")}: get the \code{WordArchive} from a parsed XML
    document  within that archive. The \code{docName} element of these
    XML documents contains the archive name and the internal element name.}
    \item{coerce}{\code{signature(from = "XMLInternalElementNode", to =
    "WordArchive")}: obtain the \code{WordArchive} object from a node in
    a parsed XML document originating from within that archive. }
    \item{coerce}{\code{signature(from = "RWordXMLDoc", to =
    "WordArchive")}: obtain the \code{WordArchive} from the parsed XML
    document that is the primary XML document within the archive. }    
	 }
}
%\references{}
\author{Duncan Temple Lang}

\seealso{
}
\examples{
showClass("WordArchive")
}
\keyword{classes}
