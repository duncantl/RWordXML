
% Generated via xml2Rd.xsl from ./tables.xml
\name{lists}
\alias{lists}
\alias{tables}
\alias{getTableNodes}
\alias{getListNodes}
\alias{getImages}
\alias{readList}
\alias{readTable}
\alias{lists,WordArchive-method}
\alias{lists,XMLInternalDocument-method}
\alias{tables,WordArchive-method}
\alias{tables,XMLInternalDocument-method}
\alias{getListNodes,WordArchive-method}
\alias{getListNodes,XMLInternalDocument-method}
\alias{getTableNodes,WordArchive-method}
\alias{getTableNodes,XMLInternalDocument-method}
\alias{getImages,WordArchive-method}
\alias{,-method}

\title{
Read lists and tables from a Word XML document
}

\description{

These functions allow us to retrieve the content of 
  and the XML nodes associated with the tables and lists within
  a Word document. 
     
}

\usage{
lists(x, nested = FALSE, ...)
tables(x, convert = readTable, ...)
getTableNodes(doc, ...)
getListNodes(doc, ...)
getImages(doc, ...)
readList(node, elFun = xmlValue, nested = FALSE, level = 0, ...)
readTable(node, as.data.frame = TRUE, colClasses = character(), header = FALSE, stringsAsFactors = FALSE, 
          elFun = cellValue, ..., rows = xmlChildren(node)[names(node) == "tr"])
}

\arguments{


\item{x}{
the WordArchive document}

\item{doc}{
the WordArchive document}

\item{node}{
the XML node that starts the list}

\item{as.data.frame}{
a logical value that controls whether the result should be structured as a data frame}

\item{colClasses}{
a character vector giving the classes for each of the columns in the table when it is converted to a data frame}

\item{header}{
a logical value that controls whether the first row of the table is treated as the header, giving the names of the column.}

\item{stringsAsFactors}{
}

\item{elFun}{
the function for processing each node in the list}

\item{nested}{
a logical value that controls whether lists within lists are returned with that structure
    or if the results are returned as a single "flat" list. }

\item{level}{
an integer value used internally for recursive calls. }

\item{rows}{
a list of the XML nodes within a table to process. This allows the caller to omit certain rows of nodes which should not be processed.}

\item{convert}{
a function that is used to process each table node in the document and convert its contents into an R object}

\item{\dots}{
additional arguments}
}

\value{


     \code{lists}
    returns either a vector of the list elements, or if nested is \code{TRUE}
     and there are nested lists, a list in which each element is either a single value or  sub-list.


     \code{tables}
    returns a list of the tables in the Word document. Each table is converted to an R object
  according to convert and so the type of the elements in the result are controlled by that function. 
  By default, these are data frames.
     

     \code{getTableNodes}
    and 
     \code{getListNodes}
    return a list of XMLInternalElementNode
objects.

}

\references{

}

\seealso{

}

  \author{Duncan Temple Lang \email{duncan@wald.ucdavis.edu}
}
 
\examples{
   #  Here we get the lists
d = wordDoc(system.file("SampleDocs", "sampleLists.docx", package = "RWordXML"))
lists(d)
         #  Here we get the lists
d = wordDoc(system.file("SampleDocs", "sampleTables.docx", package = "RWordXML"))
tables(d)
         #  Here we get the images
d = wordDoc(system.file("SampleDocs", "Images2.docx", package = "RWordXML"))
getImages(d)
        
}

\keyword{
data
}
