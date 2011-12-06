setClass("WordArchive", contains = "OOXMLArchive") #, prototype = list(classes = extends("WordArchive")))

setAs("character", "WordArchive",
      function(from)
        wordDoc(from))

wordDoc = 
function(f, create = FALSE, class = "WordArchive")
{
  createOODoc(f, create, class, system.file("templateDocs", "Empty.docx", package = "ROOXML"), "docx")
}

setMethod("hyperlinks", "WordArchive",
  # returns the content displayed in the links, i.e. the visible portion in the document
  # and the names are the targets.
  #
function(doc, comments = FALSE, ...)
{
   # Read the document
  tt = doc[[getDocument(doc)]]
  vals = hyperlinks(tt, archive = doc)

  if(comments) {
      p = getPart(doc, "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
      if(!is.na(p)) 
         vals = c(vals, hyperlinks(doc[[p]], TRUE, archive = doc))
  }

  vals
})

setMethod("hyperlinks", "XMLInternalDocument",
  # returns the content displayed in the links, i.e. the visible portion in the document
  # and the names are the targets.
  #
function(doc, comments = FALSE, archive = NULL, ...)
{

    # Get the identifiers for each w:hyperlink node so we
    # can look it up in the rels/document.xml.rels file.
  ids = xpathSApply(doc, "//w:hyperlink/@r:id")
  if(length(ids) == 0)
      return(character())

 
    # Collapse the content of the hyperlink node to text.
    # This may need to be smarter.
  vals = xpathSApply(doc, "//w:hyperlink", xmlValue)

  if(!is.null(archive)) {
     ar.el = if(comments) "word/_rels/comments.xml.rels" else "word/_rels/document.xml.rels"
     tb = getRelationships(archive, archive[[ar.el]])
     names(vals) = resolveRelationships(ids, archive, tb)
  }

  vals
})

#XXX should make this a generic.
if(!isGeneric("toc"))
  setGeneric("toc", function(file, ...) standardGeneric("toc"))

setMethod("toc", "WordArchive",
#toc = 
function(file, ...)
{
  doc = file
  if(is(doc, "WordArchive"))
   doc = doc[[getDocument(doc)]]

 tt = doc
 nodes = getNodeSet(tt, "//w:p[w:pPr/w:pStyle[contains(@w:val, 'Heading')]]")
 levels = sapply(nodes, function(x) xmlGetAttr(x[["pPr"]][["pStyle"]], "val"))
 levels = gsub("Heading", "", levels)
 structure(sapply(nodes, xmlValue), # function(x) xmlValue(x[["r"]])),
           levels = as.integer(levels),
           class = "DocumentTableOfContents")
            
# xpathSApply(tt, "//w:p[w:pPr/w:pStyle[contains(@w:val, 'Heading')]]", function(x) xmlValue(x[["r"]]))
})

print.DocumentTableOfContents =
function(x, ...)
{
  indent = sapply(attr(x, "levels"), function(i) paste(rep("   ", i-1), collapse = ""))
  cat(paste(indent, x, collapse = "\n", sep = ""), "\n")
}


getStyles =
  #
  #  This is a very simple representation of each style
  #  There are no classes at present. We just convert the XML nodes to a list for now.
  #
function(doc, useIds = FALSE)
{
  id = getPart(doc, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml") 
  xml =  doc[[ id ]] 
  ans = xpathApply( xml, "//w:style", xmlToList, namespaces = "w")
  names(ans) = sapply(ans, function(x) if(useIds) x$.attrs["styleId"] else x$name)
  ans
}

addStyle =
  #
  #
  #
function(doc, style)  
{  
  id = getPart(doc, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml") 
  xml = doc[[ id ]] 
#XXXXXXXXXXXXXXXX
  updateArchiveFiles(doc, .files = structure())
}


createPhrase =
function(text)
{
  n = newXMLNode("w:r", namespaceDefinitions = WordXMLNamespaces['w'])
  newXMLNode("w:t", parent = n, text)
  n
}


paragraph =
function(style = NA, align = NA, italic = NA, bold = NA, properties)
{  
   p = newXMLNode("w:p", namespaceDefinitions = WordXMLNamespaces["w"])
  
   props = c('w:jc' = align, 'w:pStyle' = style)  
   if(any(!is.na(props))) {
       pPr = newXMLNode("w:pPr", parent = p)
       mapply(function(id, val) {
                if(!is.na(val))
                  newXMLNode(id,  attrs = c("w:val" = val), parent = pPr)
              },  names(props), props)
    }

   if(!is.na(bold) || !is.na(italic)) {
       rPr = newXMLNode("w:rPr",  attrs = c("w:val" = align), parent = p)
       if(!is.na(bold))
          newXMLNode("w:b",  parent = rPr)
       if(!is.na(italic))
          newXMLNode("w:b",  parent = rPr)              
    }
   
   p
}




