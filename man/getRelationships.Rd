
\name{getDocument}
\alias{getDocument}
\alias{resolveRelationships}
\alias{getPart}
\alias{getRelationships}

\title{Process relationships between the different elements of an Office Open document}

\description{Office Open documents are made up of separate files that are combined together
  in various ways. They logic for how the different contents are combined is mostly controlled
  by the "relationships" which are a series of maps of part to document.
  These functions provide facilities for working with these relationships in R in order to
  make sense of an Office Open document.  
     \code{getDocument}
    gets the name of the component/file in the archive
 that corresponds to the top-level document, e.g. the word processing document or the workbook.
  
     \code{getRelationships}
    returns a data frame of relationship id, target and type triples.

     \code{resolveRelationships}
    identifies the components in the document archive
 corresponding to a particular identifier.
  
     \code{getPart}
    looks up the name of the entry in the archive corresponding to a
  conceptual part, e.g. document.
  One can index with the full name of the part or the name of the  usual file at which that part is located.
  The parts come from the PartName entries in the "[Content_Types].xml" component of the archive,
  i.e. via a call to \code{\link{contentTypes}}.
  
     \code{getPart}
    does partial matching anywhere in the string.
  
     \code{getPart}
    matches by partial name on the PartName attribute first, then by regular expression
 on the attribute PartName.
  If the part is As-Is, then we look for that literal string rather than treating the string
  as a possible regular expression.
  }

\usage{
getDocument(doc, ...)
resolveRelationships(ids, doc, relTable = getRelationships(doc), relativeTo = character())
getPart(doc, part, default = NA, stripSlash = TRUE, parts = computeParts(doc))
getRelationships(doc, rels = doc[["_rels/.rels"]])
}

\arguments{\item{doc}{the Office Open document whose properties are being queried. This is a generic OOXMLArchive object
    but typically created via a constructor function from another package which provide interfaces
   for specific types of Office Open documents..}\item{rels}{ the XML document specifying the relationships between components within the
  OO document.}\item{parts}{the collection of available parts from the document.
  If one is making repeated queries of the document, it can be convenient to compute the table of parts just once
via the (non-exported) 
\code{computeParts}
 function.}\item{\dots}{additional parameters for methods}\item{part}{a string identifying the part by file name or content type.
   Use \code{\link[base]{I}} to have the function treat it the part name as a fixed literal string rather than
regular expression.}\item{default}{the value to return if the specified part is not found/matched}\item{stripSlash}{logical value indicating whether to remove the initial / in the part name, e.g. "/word/document.xml"
  would become "word/document.xml" }\item{ids}{the identifiers that we want to resolve with respect to the relationships table in the document}\item{relTable}{the relationship table from the OO document}\item{relativeTo}{the directory to paste in front of the result}}

\value{
     \code{getDocument}
    returns a character string (vector of length 1) giving the name of the component in the
archive of the main document.
     
     \code{resolveRelationships}
    returns a character vector.
  
     \code{getPart}
    returns a character vector
  
     \code{getRelationships}
    returns a data frame with as many rows as there are relationships
   within the top-level of the document and with 3 columns/variables.
  These are the relationship identifier (id), the target component in the archive (target) and the
  type of the relationship (type).
  
  }

\references{}

\seealso{}

  \author{  \email{}
}
 
\examples{
  
  d = createOODoc(system.file("SampleDocs", "WordEg.docx", package = "ROOXML"))
  getRelationships(d)
  resolveRelationships("rId1", d)
        
  d = createOODoc(system.file("SampleDocs", "WordEg.docx", package = "ROOXML"))
  getPart(d, "numbering.xml")  # by file name
  getPart(d, I("comments+xml"))  # by content type
  
}

\keyword{documentation}
