
setGeneric("getTableNodes",
            function(doc, ...)
              standardGeneric("getTableNodes"))

setMethod("getTableNodes", "WordArchive",
            function(doc, ...)  {
               getTableNodes(doc[[getDocument(doc)]], ...)
            })

setMethod("getTableNodes", "XMLInternalDocument",
            function(doc, ...)  {              
              getNodeSet(doc, "//w:tbl", RWordXML:::WordXMLNamespaces)
            })

cellValue =
function(x)
{
  if(xmlSize(x[["p"]]) == 0)
    NA
  else
    xmlValue(x)
}

readTable =
 # Convert a table node into a data frame or matrix.
function(node, as.data.frame = TRUE, colClasses = character(), header = FALSE,
          stringsAsFactors = FALSE, elFun = cellValue, ...,
           rows = xmlChildren(node)[names(node) == "tr"])
{
  els = lapply(rows, function(x) sapply(xmlChildren(x)[names(x) == "tc"], elFun))
  ans = matrix(unlist(els), length(els), byrow = TRUE)

  if(is.logical(header) && header) {
     header = ans[1,]
     ans = ans[-1,]
   }
  
  if(as.data.frame) {
    ans = as.data.frame(ans, stringsAsFactors = stringsAsFactors)
    if(is.character(header) && length(header))
       names(ans) = header

    if(length(colClasses)) {
      ans = as.data.frame(mapply(function(x, type) {
                                   if(is.null(type))
                                      NULL
                                   else if(is.function(type))
                                      type(x)
                                   else
                                      as(x, type)
                                 },
                           ans, colClasses, SIMPLIFY = FALSE))
    }
  }

  ans   
}

if(FALSE) {
  doc = wordDoc("../ROOXML/inst/SampleDocs/WordEg.docx")  
  tbls = getTableNodes(doc)  
  o = readTable(tbls[[1]], header = c("A", "B", "C"))

  o = readTable(tbls[[2]], colClasses = c("integer", "integer", "integer"))

  all = tables(doc)

  doc = wordDoc("inst/samples/sampleTables.docx")
  all = tables(doc)
    # note the NA in all[[2]][2, 2]

  tbls = getTableNodes(doc)
  o = readTable(tbls[[1]], colClasses = c("character",  function(x) factor(x), "integer", "numeric") , header = TRUE)

    # drop the first row by handing in only the children we want. We drop the tblPr and tblGrid nodes also.
  oo = readTable(rows = xmlChildren(tbls[[1]])[-(1:3)], colClasses = c("character",  function(x) factor(x), "integer", "numeric"),
                 header = FALSE)
    # same as explicitly getting the "tr" children and dropping the first one.
  ooo = readTable(rows = (xmlChildren(tbls[[1]])[names(tbls[[1]]) == "tr"])[-1],
                  colClasses = c("character",  function(x) factor(x), "integer", "numeric"))
}



# convert = FALSE or NULL is the same as getTableNodes()
setGeneric("tables",
            function(x, convert = readTable, ...)
               standardGeneric('tables'))

setMethod("tables", "WordArchive",
            function(x, convert = readTable, ...) {
               tables(x[[getDocument(x)]], convert, ...)
            })


setMethod("tables", "XMLInternalDocument",
            function(x, convert = readTable, ...) {
              tbls = getTableNodes(x)
              if(is.null(convert))
                return(tbls)
              
              if(is.logical(convert)) {
                 if(!convert)
                   return(tbls)
                 else
                   tbls = readTable
              }
              lapply(tbls, convert, ...)
            })

