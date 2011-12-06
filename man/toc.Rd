
% Generated via xml2Rd.xsl from ./toc.xml
\name{toc}
\alias{toc}

\title{
Compute table of contents from Word docx file
}

\description{

This function returns information about the table of contents or organization
of the sections in the specified Word document.
     
}

\usage{
toc(file, ...)
}

\arguments{


\item{file}{
the docx file, either a string or an object of class WordArchive}

\item{\dots}{
additional parameters for the methods.}
}

\value{

A character vector giving the "titles" of each section found. The level
  of each section is provided via the attribute named levels.
 The class of this object is an S3 class named DocumentTableOfContents.
     
}

\references{

}

\seealso{

}

  \author{Duncan Temple Lang \email{duncan@wald.ucdavis.edu}
}
 
\examples{
   #  Sample Table of Contents
 #  This illustrates extracting the table of contents from a sample
 #  document that has several Heading 1, 2 and 3 styles.
f = system.file("SampleDocs", "toc.docx", package = "RWordXML")
doc = wordDoc(f)
tt = toc(doc)
tt
        
}

\keyword{
documentation
}
