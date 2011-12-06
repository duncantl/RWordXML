\name{toWordprocessingML}
\Rdversion{1.1}
\alias{toWordprocessingML}
\alias{toWordprocessingML,ANY-method}
\alias{toWordprocessingML,sessionInfo-method}
\alias{toWordprocessingML,table-method}
\title{Converting R output to WordprocessingML}
\description{
  This function accepts an R object and converts it into a
  WordprocessingML node which can then be inserted into the
  document.xml file within a Microsoft Word 07 archive (.docx file).
}
\usage{
toWordprocessingML(x, parent = NULL, inline = FALSE, doc = NULL, ...)
}
\arguments{
  \item{x}{the object to be converted into WordprocessingML}
  \item{parent}{the XML node, if any, which the resulting node should
    be added to.}
  \item{inline}{a logical value indicating whether the created node
    should be an inline node for insertion within a paragraph node.}
  \item{doc}{an optional XMLInternalDocument object in which the node is
    being added. This can provide the namespace definitions, etc.}
  \item{\dots}{additional arguments to be passed down to the specific
    method applied.}
}
\details{
  A number of methods for this function ship with the RWordXML package,
  including a default which simply imitates the print method for the
  object (what is shown if you type the name of the object into the R
  console).

  It is also possible to write custom methods for this function in order
  to completely control the formatting of the resulting node. See
  A New WPML Method.docx for an example of this.

  Note: We have tried to ease the process where we can, but a cursory
  understanding of the XML package and the OOXML/WordprocessingML
  specification is necessary to write new methods.
}
\value{an XML node meeting the WordprocessingML specification.
}
\references{Microsoft/ISO OOXML specification
}
\author{Duncan Temple Lang}

\seealso{

}
\examples{
require(RWordXML)

atable <- table(sample(1:10, 100, replace=TRUE))
mynode <- toWordprocessingML(atable)
mynode

}