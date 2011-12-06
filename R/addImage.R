
#
#Moved from ROOXML by Gabe  7/17/2009, and edited to add style support.
#

####################################################################
#
# Insert an image into the archive.

addImage =
  #
  #  This updates the document archive by adding the specified
  #  image file to the archive, ensuring that the Defaults are in 
  #  the [Content_Types].xml file
  #
  #  if node is an XMLInternalDocument, then add it to the end.
  #
  #
function(archive, filename, node = paragraph(align = "center"), ...,
         after = TRUE, sibling = NA, style=NA, fallback = character())
{
   newName = paste("word/media", basename(filename), sep = .Platform$file.sep)

   files = structure(lapply(filename, function(x) x), names = newName)
#   files = structure(list(filename), names = newName)

     # Add a new relationship to the document.xml.rels that identifies the new image
     # file and has an associated id so we can refer to it in the drawing node.
   rels = archive[["word/_rels/document.xml.rels"]]
     # XXX should test this is "available/free", i.e. not being used already.
   rid = paste("rId", xmlSize(xmlRoot(rels)) + seq(along = files), sep = "")

   for(i in seq(along = rid))
      addChildren(xmlRoot(rels),
                   newXMLNode("Relationship",
                              attrs = c(Id = rid[i],
                                      Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                                      Target = relativeTo(newName[i], "word"))))
   
   files[["word/_rels/document.xml.rels"]] = I(saveXML(rels))

        #XXX Need to add to the \[Content_Types].xml file and put in the
        #  entries <Default Extension="pdf" ContentType="application/pdf"/>
        # Could be smarter and just add the type of the image we are dealing with.
        # Error if we add one that is already there.
   addExtensionTypes(archive, c(pdf = "application/pdf", jpeg = "image/jpeg", png = "image/png", jpg = "image/jpeg"))

     # Add a new node
   if(!is.null(node)) {
      if(inherits(node, "XMLInternalDocument")) 
         node = xmlRoot(node)

      imgNode = createImageNode(rid, filename, ...)

      if(!is.na(style))
        addStyleToNode(node, style)

        # If sibling is a node, add imgNode as sibling of that
        # if sibling is a logical, then add imgNode as a sibling of
        # node.  Otherwise add imgNode as a child of node.
      if(inherits(sibling, "XMLInternalNode")) {
         addChildren(node, imgNode)
         addSibling(sibling, node) #XXX needs to be immediately after this sibling node, not at end.
      } else if(is.logical(sibling) && !is.na(sibling) && sibling)
        addSibling(node, imgNode)
      else
        addChildren(node, imgNode)
   }
   
#   updateArchiveFiles(archive, as(node, "XMLInternalDocument"), .files = files)
   files$document.xml = I(saveXML(as(node, "XMLInternalDocument")))
   updateArchiveFiles(archive, files) 
   if(!is.null(node))
      return(imgNode)
   else
      rid
}
