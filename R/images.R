setMethod("getImages", "WordArchive",
  #
  # Find the names of the images in the archive via the document.xml.rels file.
  #
  #XXX the root of the relative should be word/, i.e. as a prefix.
  # since this is relative word/document.xml.
  # Added now - may want to use gsub("^/", "", dirname(getDocument())). Done.
  #
function(doc, ...)
{
    # this is not a part
  part = "word/_rels/document.xml.rels"
  if(! (part %in% names(doc)))
    return(character())
 
  tt = doc[[ part ]]

  root = xmlRoot(tt)
  i = grep("image$", xmlSApply(root, xmlGetAttr, "Type"))
  prefix = gsub("^/", "", dirname(getDocument(doc)))
  structure(paste(prefix, sapply(root[i], xmlGetAttr, "Target"), sep = ""),
             names = sapply(root[i], xmlGetAttr, "Id"))
})
