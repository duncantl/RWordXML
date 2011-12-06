\name{getPackageRefs}
\alias{getPackageRefs}
\alias{getOptionRefs}
\alias{getFunctionRefs}
\alias{getClassRefs}

\title{
Extract references to R functions, classes, packages etc. from a Word document
}

\description{

 These functions allow the caller to find marked-up content in a Word document that refers
     R-related concepts such as functions, classes, packages and options.
     We can use these to ensure they are spelled correctly, identify actual R objects,
     see if they need to be renamed to match changes in code, or to programmatically
     generate, for example, bibliographies.

These functions allow one to either get the content, i.e. the text, in the document,
      or alternatively the XML nodes containing the content. We can use the latter approach to 
      then programmatically update the underlying XML document and hence the .docx file.
     
}

\usage{
getPackageRefs(doc, asNodes = FALSE, ...)
getOptionRefs(doc, asNodes = FALSE, ...)
getFunctionRefs(ar, doc = ar[[getDocument(ar)]])
getClassRefs(doc, asNodes = FALSE, ...)
}

\arguments{


\item{doc}{
the WordArchive object representing the .docx file, or
       the parsed XML document corresponding to the main document in the Word archive.}

\item{asNodes}{
a logical value }

\item{\dots}{
additional parameters passed on to  }

\item{ar}{
the WordArchive object representing the .docx file.}
}

\value{

If asNodes is \code{TRUE}, then a list of XMLInternalElementNode objects is returned.
     Otherwise, the functions return a character vector giving the content, but 
 
     \code{getFunctionRefs}
    returns a frequency table, with the table names being the function names.
     
}

\references{

}

\seealso{
\code{\link{wordDoc}}
}

  \author{Duncan Temple Lang \email{duncan@wald.ucdavis.edu}
}
 
\examples{
   #  This illustrates these functions on a document that discusses this 
 #  RWordXML package.
f = system.file("SampleDocs", "RWordDocTools.docx", package = "RWordXML")
doc = wordDoc(f)
getPackageRefs(doc)
getClassRefs(doc)
getFunctionRefs(doc)
getOptionRefs(doc)
        
}

\keyword{
programming
}
