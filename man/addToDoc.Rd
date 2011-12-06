
\name{addToDoc}
\alias{addToDoc}

\title{Add content to a document}

\description{ 
This is a generic function for which packages such as
RWordXML and RExcelXML
provide methods. These add content to an Office Open 
document. This might text such as a paragraph which is added
to the core XML document, or alternatively the addition might
be an image which is a separate file that is added to the
archive and referenced in the core document.
The intent is to abstract some of the details of inserting
new content.
     }

\usage{
addToDoc(archive, node, update = TRUE, ...)
}

\arguments{\item{archive}{the archive/office open document to which the content is being added. }\item{node}{the XML node in the main document at which the content is to be added. This identifies
  the archive element/file as well as the location.}\item{update}{a logical value indicating whether to update the archive
 on disk as well as the internal documents.}\item{\dots}{ additional parameters to be interpreted by the methods.
These include auxiliary files.}}

\value{}

\references{}

\seealso{}

  \author{}
 
\examples{
  
}

\keyword{}
