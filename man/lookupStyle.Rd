\name{lookupStyle}
\Rdversion{1.1}
\alias{lookupStyle}
\title{Looking up a Style Definition by Id
}
\description{
This function accepts style definition documents from two Word 07
archives and attempts to match a style definition to the id passed into
it. If the style definition is found in the second document, the
function either returns the style definition node, or adds the
definition to the first style definition document.
}
\usage{
lookupStyle(id, doc.styles, other.styles, add = TRUE)
}
\arguments{
  \item{id}{Id, not name, of the style whose definition should be looked
    up. A style's id is generally the name of the style, with any
    whitespace removed.
}
  \item{doc.styles}{this is the main style definition document,
    generally extracted from a wordArchive that is being actively worked
    on. If the add paramter is true, this is the document to which the
    style definition will be added.
}
  \item{other.styles}{generally a .dotm file containing additional style
    definitions that may be needed but may not already appear in the
    wordArchive being worked on.
}
  \item{add}{a logical value indicating whether the style definition
    should be added to doc.styles (if necessary).
}
}
\details{This function first attempts to find a definition for a style
  with id style.id in doc.styles. Upon failing that, it seeks the
  definition in other.styles.

  If the style is not found in either place, the function will raise an
  error.
}
\value{If the style is found ing doc.styles, the function returns NA. If
  the style is found in other.styles, the function returns either the
  style definition node, or the modified doc.styles, depending on the
  add parameter.  
}
\references{ISO/Microsoft OOXML Specification
}
\author{Gabriel Becker}

\seealso{\link{addStyleToNode}, \link{paragraph}}
\examples{

}

