
\name{getEntryType}
\alias{getEntryType}
\alias{getEntryTypeByExtension}

\title{Get the type of an element in an Office Open archive}

\description{ 
These functions allows one to determine the nature or purpose of an element in an Office Open archive.
It looks up the mapping from file name and extension to type and 
     }

\usage{
getEntryType(filename, archive, types = computeParts(archive), partialMatch = FALSE)
getEntryTypeByExtension(filename, archive, types = contentTypes(archive), partialMatch = FALSE)
}

\arguments{\item{filename}{the name of the entry in the archive, i.e. the file name of that entry.
    This can have the preceding / or not. If partialMatch is \code{TRUE} this can be any part of the file name that 
    uniquely identifies it, e.g. the base part of the name without the directory assuming that is unique within the archive.}\item{archive}{the OOXMLArchive object representing the archive.}\item{types}{the collection of mappings from file names and extensions to type names, e.g.
     "application/vnd.openxmlformats-officedocument.theme+xml" for a theme document.For 
     \code{getEntryTypeByExtension}
   , this is the content types retrieved via
  \code{\link{contentTypes}}.}\item{partialMatch}{a logical value indicating whether to treat the name of the archive
element (filename) as a partial name of an archive element.}}

\value{
A character string giving the type for the input file or extension.
     }

\references{}

\seealso{}

\examples{
   #  This shows how we can match by file name using
 #  exact matching or partial matching.
 #  If we don't use partial matching and do not have an exact match,
 #  we match by extension.
  oo = createOODoc(system.file("SampleDocs", "pics.docx", package = "ROOXML"))
  getEntryType("theme1.xml", oo)
  getEntryType("theme1.xml", oo, partial = TRUE)
        
}

\keyword{IO}
