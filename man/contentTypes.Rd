
\name{addDefaultContentType}
\alias{addDefaultContentType}
\alias{addExtensionTypes}
\alias{contentTypes}

\title{Query and update the content types of an Office Open archive}

\description{These functions allow us to query the content type of individual components/files
within an Office Open archive/document and to also add new content type information
such as extensions and the default type.
     }

\usage{
addDefaultContentType(ext, type, node)
addExtensionTypes(archive, exts, update = TRUE)
contentTypes(ar, asXML = TRUE)
}

\arguments{\item{ar,archive}{the archive object identifying the OO document.}\item{asXML}{ whether to parse the content types document as an XML document and returned
     the parsed tree.}\item{exts}{named character vector giving the extensions as names and the content type as the corresponding value.
 }\item{update}{a logical value indicating whether to update the OO archive or just
  the XML tree.}\item{ext}{a string identifying the extension for which we are adding default content type}\item{type}{a string identifying the content type for the specified extension}\item{node}{the XML node to which to add the content type}}

\value{
     \code{contentTypes}
    returns an XML document with root node
 Types. The children are 
Override and Default that specify the
content type for each "file" within the archive.
}

\references{}

\seealso{}

  \author{  \email{}
}
 
\examples{
  
}

\keyword{documentation}

\keyword{IO}
