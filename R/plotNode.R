processPlotNode =
function(node, expr, grDevice, ctr, targetDoc, Roptions = NULL, haveOutputNode = FALSE, outputNode = NULL, dir = dirname(targetDoc),
            env = new.env(), verbose = FALSE)
{

     if(FALSE && haveOutputNode) {
         # use information from what is there now
       info = getFigureInfo(outputNode, targetDoc)
         # jpeg, pdf, png, ...
       deviceName = sapply(strsplit(info$appType, "/"), function(x) x[2])       
       imageFile = paste(dir, basename(info$filename), sep = .Platform$file.sep)
     } else {
       info = list(dims = c(width = 4, height = 4))
       extension = names(grDevice)
       imageFile = paste(dir, paste("plot", ctr, ".", extension, sep = ""), sep = .Platform$file.sep)
     }  

     if(verbose)
        print(expr)


      ans = sapply(seq(along = imageFile),
                   function(i, deviceName)  {
cat("Creating", imageFile[[i]], "\n")                     
                     grFun = grDevice[[i]]
                     if("units" %in% names(formals(grFun)))
                       grFun(imageFile[i], units="in", res=72,
                             width = info$dims["width"], height = info$dims["height"])
                     else
                       grFun(imageFile[i], width = info$dims["width"],
                             height = info$dims["height"])
                     on.exit(dev.off())
                     ans = eval(expr, env)
                     if(inherits(ans, "trellis"))
                       print(ans)
                     dev.off()
                     on.exit()
                     ans
                   }, deviceName = deviceName)[[1]] #we only need the first element in the list, since they will be the same (return value of the code)
     

     if(!haveOutputNode) {
         addImage(targetDoc, imageFile, sibling = node, style=Roptions["style"])
     
     } else {  #XXXXX  need to put in both plots.
        names(imageFile) <- paste("word/media/", gsub(".*[/\\]([^/\\]*)$", "\\1",
                                                      imageFile), sep="")
        files = as.list(imageFile)
        files[[getDocument(targetDoc)]] = I(saveXML(targetDoc[[ getDocument(targetDoc) ]]))
        updateArchiveFiles(targetDoc, files)
     }     
}
