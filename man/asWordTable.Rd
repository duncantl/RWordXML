\name{asWordTable}
\Rdversion{1.1}
\alias{asWordTable}
\title{
Convert an R Table to a WordProcessingML Table
}
\description{
This function accepts an R object of class table, and creates a
corresponding WordProcessingML compliant XML node.

}
\usage{
asWordTable(tb, parent = newXMLNode("w:p", namespaceDefinitions = WordXMLNamespaces))
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{tb}{the table to be converted.
}
  \item{parent}{an XML (WordProcessingML) node which the table node
    should be added to as child. 
}
}
\value{
This function returns the parent XML node specified by parent, after the
table node has been constructed and added as a child.

}
\references{
ISO.Microsoft OOXML Specification
}
\author{Duncan Temple Lang}
\seealso{\link{toWordprocessingML}
}
\examples{

}
