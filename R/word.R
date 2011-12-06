WordXMLNamespaces =
  c(
    mv = 'urn:schemas-microsoft-com:mac:vml',
    mo = 'http://schemas.microsoft.com/office/mac/office/2008/main',
    ve = 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    o = 'urn:schemas-microsoft-com:office:office',
    r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    m = 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    v = 'urn:schemas-microsoft-com:vml',
    w10 = 'urn:schemas-microsoft-com:office:word',
    w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    wne = 'http://schemas.microsoft.com/office/word/2006/wordml',
    wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    # From Open XML Explained book (online) http://72.15.199.198/attachment/1970.ashx
    a = 'http://schemas.openxmlformats.org/drawingml/2006/main',
    pic = 'http://schemas.openxmlformats.org/drawingml/2006/picture'    
  )

if(FALSE)
  setMethod("getOOXMLPartsTable", "WordArchive", function(obj) WordprocessingMLParts)     




#XXXXXXXXXXXXXXX This version is overridden below.
if(FALSE)
asWordTable =
function(tb, parent = newXMLNode("w:p", namespaceDefinitions = WordXMLNamespaces))
{
  returnParent = missing(parent)
 
  tbl = newXMLNode("w:tbl", namespaceDefinitions = WordXMLNamespaces, parent = parent)
  grid = newXMLNode("w:tblGrid", parent = tbl)

   addRow =
     function(els) {
       r = newXMLNode("w:tr", parent = tbl)
       sapply(els, function(el) 
                      newXMLNode("w:tc", newXMLNode("w:p", createPhrase(el), namespaceDefinitions = WordXMLNamespaces), parent = r))
       r
     }
  
  if(length(rownames(tb)))
    addRow(c("", rownames(tb)))

  apply(tb, 1, addRow)
  if(returnParent)
    parent
  else
    tbl  
}

asWordTable =
function(tb, parent = newXMLNode("w:p", namespaceDefinitions = WordXMLNamespaces))
{
  returnParent = missing(parent)

  tbl = newXMLNode("w:tbl", namespaceDefinitions = WordXMLNamespaces, parent = parent)
  grid = newXMLNode("w:tblGrid", parent = tbl)

  if(length(dim(tb)) == 0)
     tb = t(as.matrix(tb))
  
  if(length(dim(tb)) == 1)
    tb = t(tb)
  
  rowNames = rownames(tb)

  newCell =
     function(el, row)
        newXMLNode("w:tc", newXMLNode("w:p", createPhrase(el), namespaceDefinitions = WordXMLNamespaces), parent = row)
  
   addRow =
     function(els, rowNum = 0) {
       r = newXMLNode("w:tr", parent = tbl, namespaceDefinitions = WordXMLNamespaces)
       if(rowNum == 0)
           newCell("", r)
       else if(length(rowNames))
          newCell(rowNames[rowNum], r)
       sapply(els, newCell, r)
       r
     }
  
  if(length(colnames(tb)))
     addRow(colnames(tb), rowNum = 0)

#  apply(tb, 1, addRow, header = FALSE)
   sapply(seq(length = nrow(tb)),
           function(i) addRow(tb[i,], i))

  
  if(returnParent)
    parent
  else
    tbl    
}


setMethod("addToDoc", c("WordArchive", "character"),
function(archive, node, update = TRUE, paragraph = FALSE, ...)           
{
  addToDoc(ar, newXMLNode("w:r",
                           newXMLNode("w:t", node,
                                         namespaceDefinitions = WordXMLNamespaces["w"]),
                           namespaceDefinitions = WordXMLNamespaces["w"]),
                update, TRUE)
})


isParagraphNode =
function(node)
{
  #XXX check the namespace.
  xmlName(node) == "p"
}

setMethod("addToDoc", c("WordArchive", "XMLInternalNode"),
function(archive, node, update = TRUE, paragraph = !isParagraphNode(node), ...)           
{
  if(paragraph) 
     node = newXMLNode("w:p", node, namespaceDefinitions = WordXMLNamespaces["w"])

  doc = xmlParse(archive[[getDocument(archive)]])
  p = getNodeSet(doc, "/w:document/w:body/w:p[last()]", WordXMLNamespaces["w"])
  if(length(p) == 0) {
    p = getNodeSet(doc, "/w:document/w:body/", WordXMLNamespaces["w"])
    addChildren(p, node)
  } else {
    p = p[[1]]
    if(xmlSize(p) == 0)
      if(isParagraphNode(node))
          #XXX or replaceNode ?  Would keep the style
         replaceNodes(p, node) # addChildren(p, kids = xmlChildren(node))
      else
         addChildren(p, node)
    else
       addSibling(p, node)    
  }
  if(update)
    archive[[getDocument(archive)]] = doc
  else
    doc
})


getStyleId =
function(node)
{
  ans = getNodeSet(node, "./w:pPr/w:pStyle/@w:val|./w:tblPr/w:tblStyle/@w:val|./w:rPr/w:rStyle/@w:val", WordXMLNamespaces["w"])
  if(length(ans))
    ans[[1]]
  else
    character()
}





getFigureInfo =
#
# This version grabs the picture information plus any information about alternative
# content (ie pdfs seem to be mac only, so they put another image file in for
# windows users).
#
# Gabe's version: function(node, ar, rid = getNodeSet(node, ".//pic:blipFill//@r:embed", WordXMLNamespaces[c("pic", "r")]))  
function(node, ar, rid = getNodeSet(node, ".//ve:Choice//@r:embed", WordXMLNamespaces[c("ve", "r")]))
{
     # Get the format of the image file, e.g pdf, jpeg, etc.
  
  filename = resolveRelationships(rid, ar, relTable = getRelationships(ar, ar[["word/_rels/document.xml.rels"]])) #XXXX 3rd arg added by gabe.
  ext = gsub(".*\\.(.*)$", "\\1", filename)
  d = contentTypes(ar)
  appType = sapply(sQuote(ext), function(x) getNodeSet(d, paste("//x:Default[@Extension =", x, "]/@ContentType"), "x"))

  
  EMU.per.inch = 914400
  extent = getNodeSet(node, ".//wp:inline/wp:extent", WordXMLNamespaces)
  dims =
    if(length(extent))
          structure(as.integer(xmlAttrs(extent[[1]])[c("cx", "cy")])/EMU.per.inch,
                     names = c("width", "height"))
    else
          c(width = NA, height = NA)
  list(filename = filename, dims = dims, appType = as.character(appType[[1]])) # or appType ?
}



createImageNode =
  #XXXX This is overridden immediately below.  Not now!
  #
  # Return an XML node and children that will display the image given in uri.
  #
function(id, filename, width = NA, height = NA, ...)
{
#cat("<createImageNode>", filename, id, "\n")  
   #XXX need to arrange to cleanup the node.
  imageTemplate = system.file("Templates", "imageTemplate.xml", package = "RWordXML")
  tmpl = xmlParse(imageTemplate, addFinalizer = TRUE)
  node = xmlRoot(tmpl)[[1]]

    # Just get the first one as the second is alternative content
  node = getNodeSet(tmpl, "//pic:blipFill/a:blip[@r:embed]", WordXMLNamespaces)
  xmlAttrs(node[[1]], append = FALSE) = c("r:embed" = id[1])
  if(length(node) > 1 && length(id) > 1)
      xmlAttrs(node[[2]], append = FALSE) = c("r:embed" = id[2])
 
  xmlRoot(tmpl)[[1]]
}


if(FALSE)
createImageNode =
function(id, filename, width = NA, height = NA, ...)
{
   #XXX Read the dimensions from filename
  if(is.na(width) && length(grep("\\.(jpg|jpeg)$", filename)) && require("rimage")) {
    d = dim(read.jpeg(filename))
    width = d[2]
    if(is.na(height)) height = d[1]
  }
  if(is.na(width)) width = 300  # was 200
  if(is.na(height)) height = 300 # was 200
  
  style = paste(c("width", "height"), c(width, height), sep = ":", collapse = ";")
  img = newXMLNode("w:pict", namespaceDefinitions = WordXMLNamespaces)
  sh = newXMLNode("v:shape", attrs = c(id = id, style = style, style = style), parent = img)  # type = "#_x0000_t75"
  newXMLNode("v:imageData", attrs = c('r:id' = id, type = "frame"), parent = sh)
  img
}



getFunctionRefs =
function(ar, doc = ar[[getDocument(ar)]])
{
   refs = xpathSApply(doc, "//w:r/w:rPr/w:rStyle[@w:val = 'Rfunc']/ancestor::w:r[1]",
                       function(x) xmlValue(x[["t"]]),
                        namespaces = c(w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"))

  refs = gsub("\\(|\\)", "", refs)
  refs = refs[ ! (refs %in% c("(", ")")) ]
  table(refs)
}


setMethod("getDocument", "WordArchive", 
function(doc, ...)
{
  getPart(doc, "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "word/document.xml")
})


setOldClass(c("RWordXMLDoc", "XMLInternalDocument"))

setAs("RWordXMLDoc", "WordArchive",
       function(from) {
         f = docName(from)
         f = strsplit(f, "::")[[1]][1]
         wordDoc(f)
       })

setMethod("[[", c("WordArchive", "character", "missing"),
            function(x, i, j, ...) {
               ans = callNextMethod()
               if(i == "/word/document.xml")
                 class(ans) = c("RWordXMLDoc", class(ans))

                 ans
            })


