\name{addStyleToNode}
\Rdversion{1.1}
\alias{addStyleToNode}
\title{Add a Formatting Style to a WordProcessingML Node
}
\description{
This function accepts an XML node which meets the WordProcessingML
specifications, as well as a character string indicating the id of a
formatting style. The node is then given that style, after any existing
style is removed.
}
\usage{
addStyleToNode(node, style.id, inline = FALSE, is.table = FALSE)
}
\arguments{
  \item{node}{the XML (WordProcessingML) node to which the style should
    be added.
}
  \item{style.id}{a character vector containing the Id (not name) of the
    style to be added. A style's id is generally it's name with any
    whitespace removed.
  }
  \item{inline}{a logical value indicating whether the style is an
    inline style (as opposed to a paragraph-based style).
  }
  \item{is.table}{a logical value indicating whether the style is a
    table style.}
}
\details{

}
\value{
Returns the node with the specified style added to it.
}
\references{
  ISO/Microsoft OOXML Specification
}
\author{Gabriel Becker}

\note{As stated above, any existing style specification on the node
  passed to this function will be removed so that the style identified
  by style.id can be added.}
\note{Currently, no check is performed to assure the style exists before
  it is added. As such, lookupStyle should be called before this
  function is used.
}



\seealso{
  \link{lookupStyle}, \link{paragraph}
}
\examples{

}

