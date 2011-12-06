
\name{hyperlinks}
\alias{hyperlinks}

\title{List the hyperlinks and text in an Office Open Document}

\description{This function examines the content of an Office Open document,
  e.g. a Word or Excel document, and finds the hyper-links within it.
  It returns the target of each links along with the text that is displayed
  for that link within the document.
  This is a generic function. The methods for the different types of documents
  are available from the corresponding package, i.e RWordXML
  and RExcelXML.
     }

\usage{
hyperlinks(doc, comments = FALSE, ...)
}

\arguments{\item{doc}{
  the OOXMLArchive document. This is typically an instance
 of WordArchive or ExcelArchive from
  the RWordXML or RExceXML package}\item{comments}{
  a logical value indicating whether to look in the content of the comments as well as the 
  main content of the document.
   }\item{\dots}{additional parameters for methods}}

\value{A character vector. The values are the text that is displayed in the document.
   The names are the target URLs of the links.
     }

\references{}

\seealso{\code{\link{createOODoc}}}

  \author{Duncan Temple Lang \email{duncan@wald.ucdavis.edu}
}
 
\examples{
   #  This shows how to get hyperlinks from a Word document, assuming
 #  RWordXML is installed.
if(require(RWordXML)) {
  d = wordDoc(system.file("SampleDocs", "hyperlinks.docx", package = "RWordXML"))
  hyperlinks(d)
  hyperlinks(d, TRUE)
}
        
}

\keyword{IO}
