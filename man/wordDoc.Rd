
% Generated via xml2Rd.xsl from ./wordDoc.xml
\name{wordDoc}
\alias{wordDoc}

\title{
Create object providing access to contents of a Word .docx file
}

\description{

This function is a constructor function for creating an R object
    that references a .docx file. 
    With this object, we can then access many different aspects of the
    Word file's content such as the text, tables, images, the comments,
    styles.  We can also update the contents of document and auxiliary
    files. 

This function returns an object which refers to the .docx file.
  As that file changes, subsequent operations on the R object retrieve
  the updated contents. This can be good or bad depending on the context
  and how the document is being updated.
     
}

\usage{
wordDoc(f, create = FALSE, class = "WordArchive")
}

\arguments{


\item{f}{
the full name of the .docx file. ~ is expanded via \code{\link[base]{path.expand}}}

\item{create}{
a logical value. If the file f does not exist and
        create is \code{TRUE} an "empty" template .docx file is created with the name f.}

\item{class}{
a character string giving the name of the S4 class to create.
     The default is typically used. However, this parameter allows other constructor functions customize the 
     class of the result. The class should derive from WordArchive,
       or at least ZipFileArchive}
}

\value{

An object of class given by the class argument.
       This is an object that typically derives from WordArchive.
     
}

\references{

}

\seealso{
\code{\link[ROOXML]{creator}}, \code{\link[ROOXML]{properties}}
}

  \author{Duncan Temple Lang \email{duncan@wald.ucdavis.edu}
}
 
\examples{
  
f = system.file("SampleDocs", "RWordDocTools.docx", package = "RWordXML")
doc = wordDoc(f)
properties(doc)
creator(doc)
created(doc)
getStyles(doc)
        
}

\keyword{

}
