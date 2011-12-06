setGeneric("getListNodes",
            function(doc, ...)
                standardGeneric("getListNodes"))

setMethod("getListNodes", "WordArchive",
           function(doc, ...) {
              getListNodes(doc[[getDocument(doc)]], ...)
           })


# XXX nested = TRUE means we try to preserve the list structure hierarchically.
#   not fully implemented yet!
setMethod("getListNodes", "XMLInternalDocument",
           function(doc, nested = FALSE, level = 0, ...) {
              xp = if(nested)
                      sprintf("//w:p[w:pPr/w:pStyle/@w:val='ListParagraph' and w:pPr/w:numPr/w:ilvl/@w:val = '%d' ]", as.integer(level))
                   else
                      "//w:p[w:pPr/w:pStyle/@w:val='ListParagraph']"

              els = getNodeSet(doc, xp)

                  # we now have to separate these as we have all the elements
                  # but we need to figure out where the end of each list
                  # occurs

#[w:p/w:pPr/w:pStyle/@w:val='ListParagraph']              
              i = !sapply(els, function(x)  {
                                   sib = getSibling(x)
                                   xmlName(sib) == "p" && "pPr" %in% names(sib) && "pStyle" %in% names(sib[["pPr"]]) &&
                                          xmlGetAttr(sib[["pPr"]][["pStyle"]], "val") == "ListParagraph"
#                                   Should be able to do this with xpath, something along ...
#                                length(getNodeSet(x, "./following-sibling::*[1][w:p/w:pPr/w:pStyle/@w:val='ListParagraph']")) == 0
                                 })
              b = cumsum(i)
              b[i] = b[i] - 1L
              if(nested) {
                levels = getListNodeLevels(, els)
                by(cbind(seq(along = els), levels), b,
                     function(x) {
                       browser() #XXXX
                     })
              } else
                tapply(seq(along = els), b, function(g) structure(els[g], class = "WordprocessingMLList"))
           })
           

readList =
  #
  # cv = wordDoc("~/Personal/CV/CVJan09.docx")
  # cv.list = getListNodes(cv)
  # pkgs = gsub(" .*", "", readList(cv.list[[1]]))
  # any(duplicated(pkgs))
  #
function(node, elFun = xmlValue, nested = FALSE, level = 0, ...)
{
   vals = sapply(node, elFun)
   if(nested) {
      # Not working yet. 
      levels = getListNodeLevels(, node)
      tapply(seq(along = node), levels, function(g) structure(vals[g], class = "WordprocessingMLList"))
   } else
    vals
}

setGeneric("lists",
            function(x, nested = FALSE, ...)
               standardGeneric('lists'))

setMethod("lists", "WordArchive",
            function(x, nested = FALSE, ...) {
               lists(x[[getDocument(x)]], nested, ...)
            })

setMethod("lists", "XMLInternalDocument",
            function(x, nested = FALSE, ...) {
              #readList(x, nested = nested, ...)
              lapply(getListNodes(x, nested = nested), readList, nested = nested)
            })



if(FALSE) {
  li = wordDoc("/Users/duncan/Books/XMLTechnologies/Rpackages/RWordXML/inst/samples/sampleLists.docx")
  z = lists(li, nested = TRUE) # not quite right yet. THe nesting lifts all the elements at the same level together
                               # not interspersed. i.e. it changes the order of the values.
}


getAllListNodes =
  # all the list nodes in the document.
function(doc)
{
  xp = "//w:p[w:pPr/w:pStyle/@w:val='ListParagraph']"
  nodes = getNodeSet(doc, xp)
  nodes
}

getListNodeLevels =
  # get the level of all the list nodes in the document.
function(doc, nodes = getAllListNodes(doc))
{
  as.integer(sapply(nodes, function(x) xmlGetAttr(x[["pPr"]][["numPr"]][["ilvl"]], "val", NA)))
}

