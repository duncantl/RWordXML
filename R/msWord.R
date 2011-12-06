# Need to update the RSID entries in word/settings.xml when we update
# the archive.

  # Create an object representing the Word document.
setGeneric("processWordDoc",
function(doc, ...)
  standardGeneric("processWordDoc"))

setMethod("processWordDoc", "character",
function(doc, ...)  
{
  if(file.exists(doc)) 
     processWordDoc(wordDoc(doc), ...)
  else 
     processWordDoc(xmlParse(doc), ...)
})

if(!is.null(getClass("WordArchive", TRUE)))
setMethod("processWordDoc", "WordArchive",
function(doc, ...)
{
  processWordDoc(doc[[getDocument(doc)]], ...)
})

setMethod("processWordDoc", "XMLInternalDocument",
function(doc, ..., styleNames = list( pPr = c("Rplot", "Rcode", "Rcodenooutput", "Rfunction", "Rfunctiondefinition", "Rtable"), rPr = "Rexpr" ))
{
    # We probably want to be smarter about line breaks, e.g. if there is a  w:proofErr, etc.
  codeNodes = getCodeNodes(doc, styleNames)
})

######################################################################

sQuote = 
function(x)
{
  sprintf("'%s'", x)
}

getCodeNodes =
  #
  # Read the document.xml file in the archive and extract the nodes with the 
  # relevant styles
  #
function(doc, styleNames = list( pPr = c("Rplot", "Rcode", "Rcodenooutput", "Rfunction", "Rfunctiondefinition", "Rtable"), rPr = "Rexpr" ),
          parse = TRUE, asNodes = FALSE, startOnly = TRUE)
{
  givenParsedXML = inherits(doc, "XMLInternalDocument")
  
  if(!givenParsedXML)
     doc = as(doc, "WordArchive")

   if(inherits(doc, "WordArchive"))
         # If we are parsing the XML document and returning the nodes, better
         # make certain that we don't allow the document to be garbage collected.
       doc = xmlParse(doc[[getDocument(doc), asXML = FALSE]], addFinalizer = TRUE) # !asNodes)
  
    # Find all the w:pPr nodes that have a style of Rplot or Rcode and then
    # get their parent which is a w:p.
  styleContainer = c(pPr = "pStyle", rPr = "rStyle")
  xpath = sapply(names(styleNames),
                  function(id) {
                   path = paste("./w:", styleContainer[id], "/@w:val", " = ", sQuote(styleNames[[id]]),
                                   collapse = " or ", sep = "")

# " and not(./parent::pPr)"
                   
                   paste("//w:", id, "[", path, "]", if(id == "rPr") "/ancestor::w:r" else "/parent::*", sep = "")
                 })
  xpath = paste(xpath, collapse = "|")

    # For Rexpr nodes, we may have content  spread across multiple nodes because of proofErr w:type='gramStart' nodes.

  codeNodes = getNodeSet(doc, xpath, WordXMLNamespaces)

  codeNodes = codeNodes[ ! sapply(codeNodes, isRexprInRcode, styleNames) ]

      # Find the inline nodes that are not separate paragraphs.
  if(startOnly) {
     i = which(sapply(codeNodes, xmlName) == "r")
     if(length(i)) {  
          # Does this deal with the expressions that span multiple elements.?
          # Yes. Does now

        w = sapply(codeNodes[i], isStartOfRexpr)
     
        if(!all(w))
           codeNodes = codeNodes [ - i [ !w ] ]
     }
  }

  if(asNodes)
    return(codeNodes)

  getCode(codeNodes, parse)
}

isStartOfRexpr =
function(node)
{
  sib = getSibling(node, FALSE)
  if(is.null(sib))
    return(TRUE)

  while(!is.null(sib) && xmlName(sib) == "proofErr") 
    sib = getSibling(sib, FALSE)    

  if(is.null(sib) ||  xmlName(sib) != "r")
      return(TRUE)

  return(length(getNodeSet(sib, "./w:rPr/w:rStyle[@w:val = 'Rexpr']", RWordXML:::WordXMLNamespaces)) == 0)
}


getLastRexprNode =
function(n)
{
  sib = n
  while(TRUE) {
    prev = sib
    sib = getSibling(sib)
     if(is.null(sib))
       return(prev)
     if(xmlName(sib) == "proofErr")
       next
     if(! (xmlName(sib) == "r" && length(getNodeSet(sib, "./w:rPr/w:rStyle[@w:val='Rexpr']", RWordXML:::WordXMLNamespaces["w"])) > 0))
       return(prev)
  }
  NULL
}


isRexprInRcode =
  # See if the node is an Rexpr within an Rcode paragraph or a derivative
  # of such a style
function(node, styleNames)  
{
  if(xmlName(node) != "r")
    return(FALSE)

  sib = getSibling(node, FALSE)
  xmlName(sib) == "pPr" && xmlSize(sib) == 2 &&
     all(names(sib) == c("pStyle", "rPr")) &&
          xmlGetAttr(sib[[1]], "val") %in% styleNames$pPr &&
          xmlGetAttr(sib[[2]][[1]], "val") %in% styleNames$rPr
}

getCodeText =
  #
  # Perhaps remove any trailing periods in R expressions that are there by accident.
  #
function(x) {
  if(xmlName(x) == "r")
    return( collapseRexpr(x))
  else if(xmlName(x) == "p") {
    vals =  xpathSApply(x, ".//w:t|.//w:br", function(x) if(xmlName(x) == "br") "\n" else xmlValue(x),
                          namespaces = WordXMLNamespaces)
    paste(c(vals, ""), collapse = "")
  } else {
    vals = xmlSApply(x, xmlValue)
    paste(c(vals, ""), collapse = "\n")
  }
}

getCode =
function(codeNodes, parse = TRUE)
{
  code = lapply(codeNodes, getCodeText)

    # This should x[["pPr"]] [[1]] or x[["rPr"]]
  names(code) = sapply(codeNodes, function(x) xmlGetAttr(x[[1]][[1]], "val"))           
  
  if(!parse)
    return(code)
  
  lapply(code, function(x) parse(text = x))
}



collapseRcode =
function(node)
{
  x = xpathSApply(node, ".//w:t|.//w:br",
                    function(node)
                     if(xmlName(node) == "br")
                       "\n"
                     else
                        xmlValue(node),
                    namespaces = WordXMLNamespaces)
  paste(x, collapse = "")
}

collapseRexpr =
function(node, checkPrevious = FALSE)
{
  if(checkPrevious && isRexprNode(getSibling(node, FALSE)))
    return(character())

  ans = xmlValue(node)
  no = getSibling(node)

  while(!is.null(no) && isRexprNode(no, FALSE)) {
    ans = c(ans, xmlValue(no))
    no = getSibling(no)
  }
  paste(ans[ ans != ""], collapse = "")
}

isRexprNode =
function(node, strict = TRUE)
{
  (!strict && xmlName(node) == "proofErr") ||
   length(getNodeSet(node, "./w:rPr/w:rStyle[@w:val = 'Rexpr']", WordXMLNamespaces)) > 0
}


#################################################################################################

setGeneric("getInlineByStyle",
  #
  # Get the references to all the R/Omegahat packages within this document.
  #
function(doc, styleNames, asNodes = FALSE, ...) 
  standardGeneric("getInlineByStyle"))


setMethod("getInlineByStyle", "WordArchive",
           function(doc, styleNames, asNodes = FALSE, ...) {
              getInlineByStyle(doc[[getDocument(doc)]], styleNames, asNodes, ...)
           })



setMethod("getInlineByStyle", "XMLInternalDocument",
function(doc, styleNames, asNodes = FALSE, ...)
{
  nodes = getRunNodesByStyle(doc, styleNames)
  if(asNodes)
     nodes
  else
     sapply(nodes, xmlValue)  
})





#################################################################################################

getPackageRefs =
 function(doc, asNodes = FALSE, ...)
      getInlineByStyle(doc, c("Rpackage", "OmegahatPackage"), asNodes, ...)

getFunctionRefs =
  function(doc, asNodes = FALSE, ...)
      getInlineByStyle(doc, "Rfunc", asNodes, ...)

getOptionRefs =
  function(doc, asNodes = FALSE, ...)
      getInlineByStyle(doc, "Roptionstyle", asNodes, ...)


getClassRefs =
  function(doc, asNodes = FALSE, ...) {
     nodes = getInlineByStyle(doc, c("RS4class", "Rclass"), asNodes = TRUE, ...)
     if(asNodes)
       nodes
     else 
        structure(sapply(nodes, xmlValue), names = sapply(nodes, getRunStyle))
  }

getRunStyle =
function(node)
{
  unlist(getNodeSet(node, ".//w:rStyle/@w:val", "w"))
}

getRunNodesByStyle =
  #
  # Find <w:r> nodes which have an rPr/rStyle node with one of the specified style names.
  #
function(doc, style)
{
  if(!inherits(style, "AsIs"))
      style = unique(c(style, unlist(sapply(style, function(x) getRoutputDescendantStyles(doc, , x)))))
  
  xpath = paste(sprintf("//w:r[./w:rPr/w:rStyle/@w:val='%s']", style), collapse = "|")
  getNodeSet(doc, xpath)
}


if(FALSE) {
  # Start of the building the picture content ourselves.
  n = newXMLNode("w:drawing", namespaceDefinitions = WordXMLNamespaces[c('w', 'wp', 'a', 'pic')])
  g = newXMLNode("a:graphic",  parent = n)
  gd = newXMLNode("a:graphicData",  parent = g, attrs = c(uri = filename))
  newXMLNode("pic:pic")
}




#addChildren(getNodeSet(tbl, ".//w:tr", namespace = WordXMLNamespaces)[[1]][[3]][[2]], createPhrase("some text in a cell")

################################################################################

StylesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"


localPNG =
pn = function(filename, width = 2000, height = 2000, units = "px", pointsize = 12, bg = "white", res = 300, ...)
         png(filename, 2000, 2000, units = units, pointsize = pointsize, bg = bg, res = res, ...)

pn = function(f, ...) png(f, 2000, 2000, res = 300)

wordDynDoc =
function(doc, out = paste(dir, "foo.docx", sep = .Platform$file.sep),
          dir = ".", env = new.env(), verbose = TRUE, removeCode = FALSE,
          nodeOp = processWordNode, force = FALSE, removeEmptyOutputNodes = TRUE,
          session.info = TRUE,
          styles = system.file("WordStyles", "RStyles.dotm", package="RWordXML"), ###### Don't want this at the top-level
#          grDevice = if(length(grep("win", targetOS, ignore.case = TRUE))) "png" else "pdf",
          grDevice = c("pdf" = pdf, "png" = pn),
          targetOS = c("windows", "mac"),
          options = list(width = width, ...), ...,
          plotExtensions = c("pdf", "png"),
          width = 70)
{

  if(file.exists(out)) {
     info = file.info(c(doc, out))
     if(info[1, "mtime"] < info[2, "mtime"] && !force)
       return(out)
  }
  
    # Copy the original document to the target given by out
  file.copy(doc, out, overwrite = TRUE)
  out = as(out, "WordArchive")

    # Note that we grab the XML document here just to make certain
    # that getCodeNodes() doesn't return the nodes and then free
    # the document. This shouldn't be a problem now!
  xdoc = out[[ getDocument(out) ]]
  
  nodes = getCodeNodes(xdoc, asNodes = TRUE)
  expr = getCode(nodes)

     # Figure out which styles are descendants of Routput and Routputtable.
     # e.g. Routputnoborder, Rplotoutput.
  out.ids = getRoutputDescendantStyles(out, target = c("Routput", "Routputtable"))


  if (is.character(styles)) {
        styleArch <- as(styles, "WordArchive")
        styles <- styleArch[[getPart(styleArch, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" )]]
   } else if (!is(styles, "XMLInternalDocument"))
       stop("Unsupported type of input for styles parameter.")

  if(length(options)) {
     if(!("width" %in% names(options())) && !is.na(width))
         options(width = width)
     
     old.opts = options()
     on.exit(options(old.opts))    
     options(options)
  }

  opts = properties(out, custom = TRUE)  #???? custom = TRUE or not? Looks like it should be TRUE.
  if( length(i <- grep("^R\\.", names(opts))) > 0) {
     if(length(options) == 0) {
       old.opts = options()
       on.exit(options(old.opts))
     }
     o = opts[i]
     names(o) = gsub("^R\\.", "", names(o))
     do.call("options", o)
  }

    #get comments for use in specifying options
   comments <- ROOXML:::commentTable2(out)
                # style
   doc.styles <- out[[getPart(out, StylesContentType)]]

  if(is.character(grDevice)) 
     grDevice = structure(lapply(grDevice, get), names = grDevice)

  if(length(names(grDevice)) == 0 || any(names(grDevice)  == ""))
     names(grDevice) = plotExtensions  
  
    # Process all of the code nodes.
  mapply(nodeOp, nodes, expr, seq(along = expr),
          MoreArgs = list(env = env, outputStyles = out.ids, dir = tempdir(),
                           verbose = verbose, target = out,
                           comments = comments, doc.styles = doc.styles, other.styles=styles, grDevice = grDevice
            #grDevice = grDevice # ? inside or outside of MoreArgs. In Gabe's code outside ???
                          ))

  if(removeCode) {
       #??? parent nodes or the nodes themselves?
       # XXX Need to get rid of the Rexpr code nodes that are not in this collection of nodes.
    removeNodes(getCodeNodes(xdoc, asNodes = TRUE, startOnly = FALSE))
  }

  if(removeEmptyOutputNodes)
    removeNodes(getNodeSet(xdoc, "//w:p[w:pPr/w:pStyle/@w:val='Remptyoutput']", WordXMLNamespaces["w"]))

  if(session.info) {
      infoNode <- toWordprocessingML(sessionInfo())
      pos <- length(xmlChildren(xdoc))
      addChildren(xmlRoot(xdoc), infoNode) # , at = pos)
      if(verbose)
        cat("Added session info.\n")
  }
  dir <- tempdir()
  tmpfilename <- paste(dir, "styles.xml", sep=.Platform$file.sep)
  saveXML(doc.styles, file = tmpfilename)  

     # and update the styles
     # 
 invisible(updateArchiveFiles(out, structure(list(I(saveXML(xdoc)), tmpfilename),
                                             names = c("word/document.xml", getPart(out, StylesContentType, stripSlash=TRUE)))))

# Mine:  invisible(updateArchiveFiles(out, xdoc))  # as(nodes[[1]], "XMLInternalDocument")))
}

if(FALSE) 
processWordNodeWithCache =
function(node, ...)
  {
    if(cacheExists(node)) {
      load(getCacheFile(node), env)
    } else {
      ans = processWordNode(node, ...)
      saveToCacheFile(objs,  node)
      
    }
  }


onWindows =
  # Determine whether we are running on Windows.
function ()
{
  length(grep("win", tolower(.Platform$OS.type))) > 0
}

processWordNode =
function(node, expr, i, targetDoc, outputStyles, dir = dirname(targetDoc),
         env = new.env(), verbose = TRUE, comments = commentTable2(targetDoc),
         doc.styles, other.styles, grDevice = if(onWindows()) "png" else "pdf")
{
#browser()  
   isPlot = isPlotNode(node)
   isTable = isTableNode(node)
   
     # Determine if the next node is an output node into which
     # we are to put the result.
  outputNode = findOutputNode(node, isPlot = isPlot, isTable= isTable)
  haveOutputNode = any(class(outputNode) == "XMLInternalNode")
    # used to be
    #   haveOutputNode = isNextSiblingRoutput(node, outputStyles)

   Roptions = getCommentOptions(node, targetDoc, comments=comments)

   if(isPlot) {
     # testing the 2 images
    return(processPlotNode(node, expr, grDevice, i, targetDoc, Roptions, haveOutputNode, outputNode, dir, env, verbose))
   }

   if(isPlot) {

     if(haveOutputNode) {
         # use information from what is there now
       info = getFigureInfo(outputNode, targetDoc)
         # jpeg, pdf, png, ...
       deviceName = sapply(strsplit(info$appType, "/"), function(x) x[2])       
       imageFile = paste(dir, basename(info$filename), sep = .Platform$file.sep)
      
     } else {
       extension = if(is.character(grDevice)) grDevice else if(onWindows()) "png" else "pdf"
       imageFile = paste(dir, paste("plot", i, ".", extension, sep = ""), sep = .Platform$file.sep)
       if(is.character(grDevice))
          grFun <- get(grDevice)
       else
          grFun = grDevice  # assuming it is a function.
       grFun(imageFile)
       on.exit(dev.off())
       if(verbose)
         cat("created", toupper(grDevice), "file", imageFile, "\n")     
     }
   }


   if(verbose)
     print(expr)


   # Evaluate the code.
#   ans = eval(expr, env)

   #  Changed by Gabe 6/29/2009 to allow for the creation of multiple image
   #files from a single code chunk, as per main and alternate images in
   #existing pictures.
   #DTL: Need an example!

   if(isPlot & haveOutputNode) {

      ans = sapply(seq(along=imageFile),
                   function(x, deviceName, imageFile, info, expr, env)  {
                     grFun = get(deviceName[x])
                     if("units" %in% names(formals(grFun)))
                       grFun(imageFile[x], units="in", res=72,
                             width = info$dims["width"], height = info$dims["height"])
                     else
                       grFun(imageFile[x], width = info$dims["width"],
                             height = info$dims["height"])
                     on.exit(dev.off())
                     ans = eval(expr, env)
                     if(inherits(ans, "trellis"))
                       print(ans)
                   }, deviceName=deviceName,
                      imageFile = imageFile,
                      info = info,
                      expr = expr,
                      env = env)[[1]] #we only need the first element in the list, since they will be the same (return value of the code)
   } else {
     ans = eval(expr, env)
     if(inherits(ans, "trellis"))
       print(ans)     
   }
   


      # If the user has indicated that they want the output suppressed,
      # by having the next paragraph be of style Remptyoutput, so be it.
      #??? Perhaps use inheritance of styles rather than a single name.
   if(haveOutputNode && getStyleId(outputNode) == "Remptyoutput")
     return(ans)

   if(isPlot) {
     if(!haveOutputNode) {
        dev.off()
        on.exit()
        addImage(targetDoc, imageFile, sibling = node, style=Roptions["style"])
     } else {
         names(imageFile) <- paste("word/media/", gsub(".*[/\\]([^/\\]*)$", "\\1",
                                                       imageFile), sep="")
         updateArchiveFiles(archive=targetDoc,doc = targetDoc[[getDocument(targetDoc)]], .files = as.list(imageFile))
     }
   } else {
      styId = getStyleId(node)
      if(grepl("nooutput$", styId))
        return(ans)
      
       if(!haveOutputNode) {
           n = toWordprocessingML(ans, inline = xmlName(node) == "r")
           if(!is.na(Roptions["style"]))
             addStyleToNode(n, Roptions["style"])
             
           insertNode(n, node)
       } else
           insertInTable(ans, outputNode)
   }   

   ans
}

insertNode =
function(n, to)
{
  if(xmlName(to) == "p") 
    return(addSibling(to, n))

  # Find the end of the code
  at = getLastRexprNode(to)
  addSibling(at, n)
}

insertInPicture =
  #
  # add the new picture to the archive
  #
  #
function(filename, picNode, ar, targetFileName = names(filename))
{
    updateArchiveFiles(ar, .files = structure(list(filename), names = targetFileName))
}

isPlotNode =
  #
  # Determines whether the given code node is for an R plot or a regular code block.
  #
function(node)
{
    length(getNodeSet(node, ".//w:pPr/w:pStyle[@w:val = 'Rplot']", WordXMLNamespaces)) > 0
}

isTableNode =
  #
  # By Gabe:
  # Determines whether the given code node is for an R table
  #
function(node)
{
  length(getNodeSet(node, ".//w:pPr/w:pStyle[@w:val = 'Rtable']", WordXMLNamespaces)) > 0
}



##################################

insertInTable =
#
# If we have a table, can we fill it in with an R object.
#
function(obj, tbl)
{
  r = nrow(obj)

  rowNodes = getNodeSet(tbl, "./w:tr", "w")
  num.format = paste("%.", getOption("digits"), "f", sep = "")
  for(i in seq(length = r)) {
    row = rowNodes[[i]]
    tc = xmlChildren(row)[names(row) == "tc"]
    sapply(seq(along = obj[i,]),
           function(j) {
               # Not just the first w:t, but need to try to shape the
               # R value to this
              text = getNodeSet(tc[[j]], ".//w:t", "w")[[1]]
              val = obj[i,j]
              if(is.numeric(val))
                val = sprintf(num.format, val)
              xmlValue(text) = val
           })

  }
  tbl
}



#####################

#
# The idea is to look at the next after the code node and see if it 
# is a paragraph with a style that "is based on" the Routput style
# either directly on indirectly.
#

isNextSiblingRoutput =
  #
  #
  #
function(node, outputStyleIds)
{
  nextNode = getSibling(node)

  id = getStyleId(nextNode)
  if(length(id) == 0)
    return(FALSE)

  id %in% outputStyleIds
}


getRoutputDescendantStyles =
  #
  # Finds the ids of the styles which are "based on"  Routput
  # We would compute this when we read the document - just once -
  # and then use this to find if the next sibling node after a 
  # code node is an R output node.
  #
function(doc, styles = getStyles(as(doc, "WordArchive"), TRUE), target = "Routput")
{
  ss = styles
  ans = target
  while(length(target) > 0 && length(ss)) {
     bon = sapply(ss, function(x) if("basedOn" %in% names(x)) as.character(x$basedOn) else "")
     i = bon %in% target
     w = unique( names(ss)[ i ])
     ans = c(ans, w)
     target = w
  }
  ans
}

if(!isGeneric("source"))
  setGeneric("source",   function (file, local = FALSE, echo = verbose, print.eval = echo, 
             verbose = getOption("verbose"), prompt.echo = getOption("prompt"), 
             max.deparse.length = 150, chdir = FALSE, encoding = getOption("encoding"), 
             continue.echo = getOption("continue"), skip.echo = 0, keep.source = getOption("keep.source"))
                  standardGeneric("source"))

setMethod("source", "WordArchive",
  function (file, local = FALSE, echo = verbose, print.eval = echo, 
             verbose = getOption("verbose"), prompt.echo = getOption("prompt"), 
             max.deparse.length = 150, chdir = FALSE, encoding = getOption("encoding"), 
             continue.echo = getOption("continue"), skip.echo = 0, keep.source = getOption("keep.source"))
{
   styles = getRoutputDescendantStyles(file) 
   xdoc = file[[ getDocument(file) ]]
   nodes = getCodeNodes(xdoc, asNodes = TRUE)
   expr = getCode(nodes)

   env = if(is.logical(local))
           if(local) sys.call(sys.nframe()) else globalenv()
         else
           local
   
   invisible(lapply(expr, evalNode, verbose = verbose, env = env))
})

evalNode =
function(expr, env, verbose = FALSE)
{
  if(length(expr) == 0)
    return(NULL)
  
  if(verbose) {
     cat("***** evaluating\n")
     print(expr)
     cat("\n*****\n")     
  }
  eval(expr, env)
}


