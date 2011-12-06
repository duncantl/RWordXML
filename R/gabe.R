findOutputNode =
  #
  # Gabe, 7/27/2009
  # Attempt to figure out if there is an output node corresponding to the code block
  # Returns TRUE if there is (seems to be) and FALSE if no such node was found.
  # Unlike the version above, this treats all code without the Rplot style
  # as code that potentially generates a table. ie it looks for a table
function(node, isPlot = isPlotNode(node)# , isTable = isTableNode(node), archive,
#         comments= commentTable2(archive),
#         commentOptions = getCommentOptions(node, comments=comments)
         )
{
  {
    siblingNode <- getSibling(node)

    if (isPlot)
    {
      length(getNodeSet(siblingNode, ".//w:r//w:drawing|.//w:r//w:pict",
                        WordXMLNamespaces)) > 0 
    } else {
      xmlName(siblingNode) == "tbl"
      
      
    }
  }
}


findOutputNode =
  #
  # Gabe, 7/27/2009
  # Attempt to figure out if there is an output node corresponding to the code block
  # Returns the output node, if one is found, and NA if not.
  # This version assumes that the Rtable style has been implemented and is being used
  # to indicate code that generates tables (as Rplot does for plots), though only
  # if the default value of isTable is used. this could be changed to isTable = !isPlot
  # if the Rtable style does not end up getting implemented, though this would
  # cause the issues/problems of not knowing whether a table is needed until
  # after evaluating the code.
  #
  # This seems like the better solution, as it is unclear how we can easily tell
  # whether to look for a table or not if the code is not marked as Rtable. Also
  # users will already be used to the idea because of the necessity of Rplot.
function(node, isPlot = isPlotNode(node), isTable = isTableNode(node),
         searchLength = 1#, archive,
#         comments= commentTable2(archive),
#         commentOptions = getCommentOptions(node, comments=comments)
         )
{
  if (searchLength < 1 || searchLength != as.integer(searchLength))
    stop("Invalid searchLength. searchLength must be a positive integer")
  
  curNode = node
  for (i in 1:searchLength)
  {
    siblingNode = getSibling(curNode)
    #If there are no siblings left to check, return NA.
    if (is.null(siblingNode))
      return(NA)
    if(isPlot || isTable)
    {
      #if isTable is TRUE and we find a table, return siblingNode (the output node)
      if (isTable && xmlName(siblingNode) == "tbl")
        return(siblingNode)

      #if isPlot is TRUE and we find an image, return siblingNode (the output node)
      if (isPlot &&
          length(getNodeSet(siblingNode, ".//w:r//w:drawing|.//w:r//w:pict",
                            WordXMLNamespaces)) >0)
        return(siblingNode)
    }
    #advance the current node to continue the search
    curNode = siblingNode
  }
  #if nothing was found, return NA
  NA
}


getCommentOptions =
  #
  # Gabe 7/27/2009
  # Compile options specified in comments attached to a particular code node
  #3
function(node, archive, comments=commentTable2(archive))
{
  if (missing(archive) && missing(comments))
    stop("archive and comments are both missing. At least one must be specified.")

  commentIds <- as.numeric(xpathApply(node, ".//w:commentRangeStart/@w:id",
                                      namespaces=WordXMLNamespaces))

  #need to add support for styles descended from Roptions
  commentRows <- which(comments$id==commentIds, comments$style=="Roptions")

  if (length(commentRows))
  {
    Roptions <- comments[commentRows,"value"]
    Roptions <- paste(Roptions, sep="") #get it all in a single string
    splitOpt <- unlist(strsplit(Roptions, "[=,]"))
    ind <- seq(2, length.out=length(splitOpt)/2, by=2)
    Roptions <- splitOpt[ind]
    names(Roptions) <- splitOpt[-ind]
  } else Roptions <- character()

  Roptions
}




wordSource =
  #
  # This parses docx file  and runs the Code in the R session, but does not
  # output a new docx file.
  #
  # Currently, things that are meant to just print out to the console, ie
  # "X + 5" don't work. Should talk to Duncan about it.
  #
  # Also, not sure I like the asterisks in verbose mode.
  #

function(f, verbose=TRUE, env=parent.frame())
{
  force(env)
  arch = as(f, "WordArchive")
  doc = arch[[getDocument(arch)]]

  code = getCodeNodes(doc)

  sapply(code, evalNode,verbose=verbose, env=env )
  invisible(NULL)
}


#
#This appears to be a helper function written solely to assist addImage,
#which I migrated from ROOXML to this package. Thus I have imported this
#function as well. Currently there is still a copy in my version of ROOXML
#but as soon as I confirm with Duncan that it's not used for anything else
#that will be removed

#
# This probably belongs in ROOXML.
#
relativeTo =
function(filename, dir, sep = .Platform$file.sep)
{
  els = strsplit(filename, sep)[[1]]
  if(els[1] == dir)
    els = els[-1]
  paste(els, collapse = sep)
}
