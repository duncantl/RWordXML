\name{paragraph}
\Rdversion{1.1}
\alias{paragraph}
\title{Create a WordProcesssingML Paragraph Node
}
\description{
This function creates a paragraph node (according to the
OOXML/WordProcessingML Specification) with specified properties.
}
\usage{
paragraph(style = NA, align = NA, italic = NA, bold = NA, properties)
}
\arguments{
  \item{style}{a character vector containing the id of a paragraph-based
    style which should be added to the created paragraph node.
  }
  \item{align}{a character vector containing the desired alignment of
    text and objects in the paragraph.
}
  \item{italic}{a logical value indicating whether text in the paragraph
    should be italicized.
}
  \item{bold}{a logical value indicating whether text in the paragraph
    should be in bold.
}
  \item{properties}{currently ignored.
}
}
\value{This function returns an XML node which defines a paragraph
  according to the OOXML/WordProcessingML specification.
}
\references{ISO/Microsoft OOXML Specification
}
\author{Duncan Temple Lang
}
\note{Currently, no check is made to ensure that the style indicated in
  the style argument exists. It is advisable to call lookupStyle before
  calling this function with a non-NA style parameter.}

\seealso{
  \link{lookupStyle}, \link{addStyleToNode}
}
\examples{

}

