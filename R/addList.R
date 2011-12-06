#
# The functions here allow one to add a list to a Word document.
# They take care of handling the relationships within the document
# to connect the list with a numbering id.  It creates the word/numbering.xml
# file if necessary.


toWMLText  =
function(x, parent = NULL)
{
  newXMLNode("w:t", format(x), parent = parent)
}

getNextListId =
  #
  # Make up the next id, either sequentially or randomly.
  #
function(doc, xdoc = doc[[ ROOXML:::getPart(doc, "numbering", "word/numbering.xml") ]], random = FALSE)
{
   ids = unlist(getNodeSet(xdoc, "/*/w:abstractNum/@w:abstractNumId", RWordXML:::WordXMLNamespaces["w"]))
   ans = as.character(length(ids) + 1)
   if(random || ans %in% ids) 
      ans = paste(sample(c(0:9, LETTERS[1:6]), 8, replace = TRUE), collapse = "")

   ans  
}

addList =
  #
  # Main entry point for adding a list to a Word document.
  #
  #
function(obj, doc, node = NULL, id = NA, lvl = 0, ..., inline = TRUE, elFun = toWordprocessingML)
{
  other = addListMetaData(doc, id, lvl = lvl, ...)
  
  els = makeList(obj, other$id, node, lvl = lvl, inline = inline, elFun = elFun)

  if(is.null(node)) {
     xmlDoc = doc[[getDocument(doc)]]
     root = xmlRoot(xmlDoc)
     body = root[["body"]]
     addChildren(body, kids = els, at = xmlSize(body) - 1)
  } else
    xmlDoc = as(node, "XMLInternalDocument")
  
  # list(meta = other, doc = as(node, "XMLInternalDocument"))
  updateArchiveFiles(doc, xmlDoc, .files = other$files)
}

addListMetaData =
  #
  # Update or create the word/numbering.xml file.
  #
function(doc, id = NA, lvl = 0, ..., xdoc = doc[[ ROOXML:::getPart(doc, "numbering", "word/numbering.xml") ]])
{
  if(is.null(xdoc))
    xdoc = emptyNumberingDoc()
    
  if(is.na(id))
    id = getNextListId(doc, xdoc)

  makeListMetaNode(id, lvl = lvl, ..., parent = xmlRoot(xdoc))

  list(files = list("word/numbering.xml" = xdoc), id = id)
}

makeListMetaNode =
  #
  # Create and add the w:abstractNum node.
  #
function(id, lvl = 0, ..., parent = NULL)
{
  abs = newXMLNode("w:abstractNum", attrs = c("w:abstractNumId" = id), parent = parent)

  lvl = newXMLNode("w:lvl", attrs = c("w:ilvl" = lvl), parent = abs)
  
  args = list(...)
  if(length(args))
    mapply(function(id, val) {
             newXMLNode(id, attrs = c("w:val" = val), parent = lvl)
           },
           paste("w", names(args), sep = ":"),
           args)
  
  invisible(abs)
}

emptyNumberingDoc =
  #
  # Create an empty numbering document.
  #
function()
{
  root = newXMLNode("w:numbering", namespaceDefinitions = RWordXML:::WordXMLNamespaces["w"])
  newXMLDoc(node = root)
}

makeList =
  #
  # Make the nodes for a list.
  #
  #
function(els, id = NA, parent = NULL, elFun = toWordprocessingML, lvl = 0, inline = TRUE)
{
  lapply(els, makeListEl, id, parent, elFun, lvl, inline)
}

makeListEl =
  # 
  # Make an individual node for a list.
  #
function(el, id, parent, elFun, lvl, inline = TRUE)
{  
  p = newXMLNode("w:p", namespaceDefinitions = RWordXML:::WordXMLNamespaces["w"], parent = parent)
  pPr = newXMLNode("w:pPr", parent = p)
    newXMLNode("w:pStyle", attrs = c("w:val" = "ListParagraph"), parent = pPr)
  
  numPr = newXMLNode("w:numPr", parent = pPr)
     newXMLNode("w:ilvl", attrs = c("w:val" = lvl), parent = numPr)
     newXMLNode("w:numId", attrs = c("w:val" = id), parent = numPr)
# if(inline)
#    pp = newXMLNode("w:r", parent = p)
# else
     pp = p

  elFun(el, parent = pp, inline = inline)

  p
}

