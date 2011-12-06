\name{wordDynDoc}
\alias{wordDynDoc}
\title{Process contents of a dynamic Word '07 document, evaluating R code}
\description{
 This function provides the dynamic processing facilities for a Word
 document.
 It reads the text of a Word '07 document and evaluates the
 R code segments identified by various styles within the document.
 It inserts the converted output from the expressions
 into the document.  It uses any existing formats following the
 R code elements and attempts to insert the R object resulting
 from the code into that format.  This means that the author can
 customize the appearance of the results using regular Word
 facilities, e.g. creating a table to hold the result and
 interactively changing the column widths, the background,
 or inserting a template image which is replaced by an
 R graphic, allowing the user to specify the size of the graphic and any
 other formatting (e.g. alignment).

 This function processes a .docx file which is a Zip archive.
 It uses the \code{Rcompression} to read that file and currently
 requires the zip executable/command to be available to insert
 elements into the new archive.

 This function does not require Microsoft Office or Word to be
 installed. It does not utilize Word to create the new document.  It
 parses the XML documents within the Zip archive and updates that XML
 document and the contents of the Zip file.  The result can be displayed
 in Word 2007 on any machine that supports it.

 More work can be done to make the insertion of R objects into
 existing Word elements more flexible.
 For example, if a matrix is to be inserted into a table with too few
 columns, the table should be extended, using the R recycling rule.
}
\usage{
wordDynDoc(doc, out = paste(dir, "foo.docx", sep = .Platform$file.sep),
            dir = ".", env = new.env(), verbose = TRUE,
           removeCode = FALSE, nodeOp = processWordNode,
           force = FALSE, removeEmptyOutputNodes = TRUE,
           session.info = TRUE,
           styles = system.file("WordStyles", "RStyles.dotm", package = "RWordXML"), 
           grDevice = c(pdf = pdf, png = pn),
           targetOS = c("windows", "mac"),
           options = list(width = width, ...),
           ...,
           plotExtensions = c("pdf", "png"),
           width = 70
          )
}

\arguments{
  \item{doc}{the file name of the Word document (a docx file).}
  \item{out}{the name of the new file to be created by processing the
    contents of \code{doc}.}
  \item{dir}{directory in which to write the resulting file.
    Note that graphics files are written to a temporary directory
    and are inserted into the Word archive, so are essentially anonymous
    and are not retained as explicit files after the new document
    is created.
   }
  \item{env}{an environment in which the R expressions are evaluated.
    This provides the caller with the opportunity of customizing the
    evaluation, e.g. initializing variables that are referenced in the
    code within the document.}
  \item{verbose}{a logical value that controls whether information is
    displayed on the console as each code segment is processed.
    This information includes the creation of graphics devices.
  }
  \item{force}{a logical value that controls whether the document is
  processed even if the output file is already more recent than the
  input file.}
   \item{removeCode}{a logical value that controls whether the
     R code content (XML nodes) within the source document are
     removed from the output (\code{TRUE}) or left in (\code{FALSE}).
   }
   \item{nodeOp}{the function which is called to process each R code node
     within the Word document. This is called with the
    XML node, the R expression, the index/number of the expression being
    processed, the target document as a \code{WordArchive} object,
    a character vector giving the names of all the Word styles
    identifying R output nodes,
    the output directory, the environment in which to evaluate the R
    expressions,
    and a logical value indicating whether to emit information as the
     evaluation progresses.
   }
 \item{removeEmptyOutputNodes}{ a logical value controlling whether we
  remove empty code nodes that might have been left by the author.}
 \item{session.info}{a logical value indicating whether information
   about the R session should be inserted at the end of the output
   document.}
 \item{styles}{a character string indicating a Word '07 document
   (generally a macro-enabled template, or .dotm, file) which contains
   any additional styles which may be desired for the output document,
   but which are not included in the input document because they were
   not used. The RWordXML package ships with a file called Rstyles.dotm
   which is the default for this argument. This should only be changed
   if user-created styles are to be used.}
 \item{options}{a list of global R options that are set when for the
   duration of the processing of the code in the Word document.}
 \item{width}{the value to set the R global \code{width} option
  controlling the output on the screen.}
 \item{grDevice}{a list of functions or a character string containing the name of the functions
   used to create the desired graphics device. If given as a list, the
  list should have names that identify the type of plot the function
  will create, and these will be used as extensions for the generated
  file names. If the list does not have names, one can specify these
  extenion names and plot types   via the \code{plotExtensions}
  parameter. %The default depends on the targetOS argument.
}
 \item{plotExtensions}{names of the types of plots created by the
  graphics devices in \code{grDevice}. These are used as extensions for
  file names of the plots generated during the processing of graphics
  code. This argument is only used if \code{grDevice} is a list of
  functions with no names.
  }
 \item{targetOS}{(not necessarily used now) a character vector indicating which operating
   system(s) the resulting document is intended to be viewed on. This
   information is then used to choose the graphics device that will be
   used to generate image files. Valid entries are "windows" and "mac".
   This argument will be ignored if the graphics device is specified
   directly using the grDevice argument.}
 \item{\dots}{any "loose" name = value pairs that are to be set as part
  of a call to \code{\link[base]{options}} that will persist for the
  duration of processing the code in the document.
  Code within the document can change these options.
  Also the original options from before the call are reset. So these are
  temporary options. One can also specify these values as part of a list
  via the \code{options} parameter.}
}
\details{
 There are two ways to dictate the format/appearance of output inserted
 into the output document.  %XXXX
 
 Options regarding how the output from individual blocks of code is
 inserted into the document can be specified by attaching comments to
 the code from within Word '07. See CommentOptions.docx.
}
\value{
  An object of class \code{WordArchive}  identifying the newly created file.
}
\references{
 Microsoft Office 2007
}
\author{Duncan Temple Lang}
\seealso{
 An alternative format and approach is provided by
 \code{\link[XDynDocs]{dynDoc}} in the \code{XDynDocs} package.
}
\examples{
 f = system.file("SampleDocs", "WordEg.docx", package = "RWordXML")
 wordDynDoc(f, "bar.docx", force = TRUE)
}
\keyword{IO}
\keyword{programming}
\concept{dynamic documents}
\concept{reproducible research}

