findNodesByStyle =
function(doc, styles, paragraph = TRUE)
{
  if(is(doc, "WordArchive"))
    xmlParse(doc [[ getDocument(doc) ]], addFinalizer = TRUE) # XXXX should be okay to add finalizer now.
  if(paragraph)
    xpath = paste("//w:pPr/w:pStyle[", paste("@w:val =", sQuote(styles), collapse = " or "), "]/..")

  getNodeSet(doc, xpath, WordXMLNamespaces)
}

getSections =
function(doc, levels = 1:4, asNodes = FALSE)
{
  nodes = findNodesByStyle(doc,  paste("Heading", levels, sep = ""))
  txt = sapply(nodes, function(x) xmlValue(xmlParent(x)))
  if(asNodes) {
    names(nodes) = txt
    nodes
  } else
    txt
}
