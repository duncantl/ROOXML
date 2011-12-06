
\name{properties}
\alias{properties}
\alias{keywords}
\alias{Title}
\alias{property}
\alias{modified}
\alias{description}
\alias{category}
\alias{created}
\alias{creator}
\alias{revision}
\alias{subject}

\title{Query properties of an Office Open Document.}

\description{ These functions allows us to query properties of the 
   global Office Open document such as an Excel or Word document.  
  These properties are items such as the name of the author, the date the document was created
 and when it was modified, generic comments, keywords, the category of the document.
 We can also retrieve custom properties, i.e. those that are explicitly added by the 
   author(s) that are not part of the general collection of properties.
  These are more unusual properties and also free-form properties that 
  authors can create themselves.
     }

\usage{
properties(doc, ..., tt = doc[[if (custom) "docProps/custom.xml" else "docProps/core.xml"]], 
           .els = as.character(unlist(list(...))), custom = FALSE)
keywords(doc, sep = "[,;]")
Title(doc)
property(doc, el, tt = doc[["docProps/core.xml"]])
modified(doc)
description(doc)
category(doc)
created(doc)
creator(doc)
revision(doc)
subject(doc)
}

\arguments{\item{doc}{the Office Open document whose properties are being queried. This is a generic OOXMLArchive object
    but typically created via a constructor function from another package which provide interfaces
   for specific types of Office Open documents..}\item{\dots}{character strings specifying which properties are being requested.
   If none are specified, all the available properties are returned.}\item{tt}{the document containing the properties being queried. If one is querying numerous properties
    individually, it may be convenient to fetch this document once and pass it in several calls. Typically this argument will be omitted.}\item{.els}{an alternative way to specify the sub-collection of desired properties.
   This is convenient when one has the names as a character vector and not as individual names in the call.}\item{custom}{a logical value indicating whether we want
to work with the custom properties that the authors of the document may have added
explicitly to the document, including free-form properties.
\code{TRUE} indicates that we should use the custom properties; \code{FALSE} indicates
using the regular document properties.
}\item{sep}{
a regular expression or literal string that is used to 
split the entries in the keyword string. This is passed
to \code{\link[base]{strsplit}}.
}\item{el}{
a character string giving the name of the property to be retrieved.
}}

\value{A list with the property values.
     }

\references{
   Office Open
   }

\seealso{\code{\link[RWordXML]{wordDoc}} and \code{\link[RExcelXML]{excelDoc}} in 
   the RWordXML and RExcelXML packages respectively.
   }

  \author{Duncan Temple Lang \email{duncan@wald.ucdavis.edu}
}
 
\examples{
   #  The following shows how we can query an Excel document.
         file = system.file("SampleDocs", "Workbook2.xlsx", package = "ROOXML")
         doc = createOODoc(file, class = "OOXMLArchive")

         properties(doc)
         names(properties(doc))

         creator(doc)
         created(doc)
         modified(doc)
         properties(doc, "lastModifiedBy")
        
}

\keyword{IO}
