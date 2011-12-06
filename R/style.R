lookupStyle=
#
#This function will look for style definitions, first in the document
# styles.xml file as it exists now, and then in the styles defined in an
#additional xml document. It will raise an error if the style (identified
#by id) is not found in either place. If add is TRUE then the style
#definition node from other.styles will be added to the doc.styles
#document if necessary.
#

function(id, doc.styles, other.styles, add=TRUE)
{
  stylenode <- getNodeSet(doc.styles, paste("//w:style[@w:styleId=", id, "]", sep=""), namespaces=WordXMLNamespaces)
  if (length(stylenode))
    return(NA)

  stylenode <- getNodeSet(other.styles, paste("//w:style[@w:styleId='", id, "']", sep=""), namespaces=WordXMLNamespaces)

  if (!length(stylenode))
    stop(paste("Style", id, "not found."))

  if (add)
    invisible(addChildren(xmlRoot(doc.styles), stylenode[[1]]))
  else
    stylenode[[1]]
}


addStyleToNode =
function(node, style.id, inline=FALSE, is.table = FALSE)
{
  if (is.table)
    char = "tbl"
  else
  {
    char = if(inline) "r" else "p"
  }
  PrNode = xmlChildren(node)[paste(char, "Pr", sep="")]
  if(!length(PrNode))
    PrNode= newXMLNode(paste("w:", char, "Pr", sep=""),
      parent=node, at=1)
  else
    PrNode = PrNode[[1]]

  styleNode <- xmlChildren(PrNode)[paste(char, "Style", sep="")]
  
  if(length(styleNode[[1]]))
    removeNodes(PrNode, styleNode[[1]])

  styleNode = newXMLNode(paste("w:", char, "Style", sep=""), attrs=list("w:val"= style.id),
              namespaceDefinitions=WordXMLNamespaces)
  addChildren(PrNode, styleNode)
  
  invisible(node)
  
}
