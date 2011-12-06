setGeneric("toWordprocessingML",
           function(x, parent = NULL, inline = FALSE, doc = NULL, ...)
             standardGeneric("toWordprocessingML"))


setMethod("toWordprocessingML", "table",
function(x, parent = newXMLNode("w:p", namespaceDefinitions = WordXMLNamespaces["w"]), inline = FALSE, doc = NULL, ...)
{
   asWordTable(x,parent)
})


setMethod("toWordprocessingML", "ANY",
  #
  # Print the object as it appears in R and put it in a "verbatim" style"
  #
function(x, parent = newXMLNode(if(inline) "w:r" else "w:p", namespaceDefinitions = WordXMLNamespaces["w"]),
          inline = FALSE, doc = NULL,  style=if(inline) "Rinlineoutput" else "Routput", ...)
{
  z = if(inline && is.vector(x) && length(x) == 1)
          format(x)
      else
          capture.output(print(x))

    # Add the no proof
  if(inline)
    newXMLNode("w:rPr", parent = parent, namespaceDefinitions = WordXMLNamespaces["w"],
                newXMLNode("w:rStyle", attrs = c("w:val" = style), #XXX use the id not the name!
                            namespaceDefinitions = WordXMLNamespaces["w"]))    
  else
    newXMLNode("w:pPr", parent = parent, namespaceDefinitions = WordXMLNamespaces["w"],
                newXMLNode("w:pStyle", attrs = c("w:val" = style), #XXX use the id not the name!
                            namespaceDefinitions = WordXMLNamespaces["w"]))
  sapply(seq(along = z),
          function(line) {
# Don't use the code in comments. Testing adding lists.            
#        if(inline) {
#          if(line > 1)
#             newXMLNode("w:br", namespaceDefinitions = WordXMLNamespaces["w"], parent = parent)
#          newXMLNode("w:t", z[line], namespaceDefinitions = WordXMLNamespaces["w"], parent = parent)
#        } else
                newXMLNode("w:r", if(line > 1) newXMLNode("w:br", namespaceDefinitions = WordXMLNamespaces["w"]),
                                  newXMLNode("w:t", z[line], namespaceDefinitions = WordXMLNamespaces["w"]), parent = parent)
          })
  invisible(parent)
})


setOldClass("sessionInfo")
setMethod("toWordprocessingML", "sessionInfo",
function(x, parent = newXMLNode(if(inline) "w:r" else "w:p", namespaceDefinitions = WordXMLNamespaces["w"]),
          inline = FALSE, doc = NULL,  ...)
{
  ans = callNextMethod()
  setStyle(ans, I("RSessionInfo"), doc = doc)
  ans
})

setStyle =
function(node, name, check = !inherits(name, "AsIs"), force = FALSE,
          doc = as(node, "WordArchive"))
{
  id = name
  if(is.null(doc))
    id = gsub(" ", "", id)
  
  if(check && !is.null(doc)) {
     styles = getStyles(doc)
     if(!(id %in% names(styles))) {
        if(!checkStyleId(id, styles) && !force) 
            stop("Cannot find style ", id)

     } else {
        id = styles[[id]]$.attrs["styleId"]
     }
  }
  
  xmlAttrs(node[[1]][[1]]) <- c('w:val' = id)
}

setAs("XMLInternalElementNode", "WordArchive",
      function(from) {
          doc = as(from, "XMLInternalDocument")
          as(doc, "WordArchive")
      })

setAs("XMLInternalDocument", "WordArchive",
      function(from) {
          fname = strsplit(docName(from), "::")[[1]][1]
          wordDoc(fname)
      })  


checkStyleId =
  #
  # checks the specified id is a valid style  id.
  #
function(id, styles)
{
   w = sapply(styles, function(x) x$.attrs["styleId"] == id)
   any(w)
}
