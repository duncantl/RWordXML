\name{addList}
\alias{addList}
\title{Add a list to a Word document}
\description{
  This function adds a Word list to a word processing document.
  It takes the R object and attempts to represent it as a list.

  In the future, we should make this generic. Howver, the
  \code{makeList} function is the ont that creates the XML
  representation of the R object. So making this generic is
  more releated to \code{doc} and \code{node} and being able to add
  to a parsed XML document or to  a particular place in the document,
  rather than just the archive. However, one can do this via \code{node}
  now.
}
\usage{
addList(obj, doc, node = xmlRoot(doc[[getDocument(doc)]]), id = NA, lvl = 0, ..., inline = TRUE)
}
\arguments{
  \item{obj}{the R object which is to be displayed as a list in the
  Word document.}
  \item{doc}{the Word document, i.e. of class \code{\link{WordArchive-class}}}
  \item{node}{the node in the primary XML document within the Word
  archive at which to put this new node. The new list is added as a
  child of this node.}
  \item{id}{the identifier for the list that is used in the Word meta
  data to index this list.}
  \item{lvl}{the level of the list, i.e. the level of nesting.}
  \item{\dots}{additional parameters that are passed to the
  \code{addListMetaData} function.}
  \item{inline}{a logical value passed to \code{makeList} and which
  controls whether the content is add as a child of a \code{w:r} node or
   directly within a \code{w:p} node.}
}
\value{
  The side-effect of updating the archive is the primary purpose of this
  function.
  
  If we add methods for this function to update the parsed XML document,
  the archive will not be updated directly.
}
\author{
Duncan Temple Lang
}
\seealso{

}
\examples{
 doc = wordDoc("doesNotExist.docx", create = TRUE)
 addList(c("Duncan", "John", "Naras"), doc)
 lists(doc) # doesn't work yet, but the content is display.
}
\keyword{programming}
\keyword{documentation}

