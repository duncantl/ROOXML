
\name{setUpdated}
\alias{setUpdated}
\alias{creator<-}
\alias{setProperty}

\title{Set one or more global properties of an Office Open document}

\description{These functions allow us to modify and update the properties of a document
     to set things like the author/creator, date modified, etc.
    
     \code{setProperty}
    allows us to set general properties by name.
     
     \code{creator<-}
    and 
     \code{setUpdated}
    are convenience functions which set 
     specific properties.
These functions update the external archive if doc  is
   descends from OOXMLArchive.}

\usage{
setUpdated(doc, date = Sys.time(), update = TRUE)
'creator<-'(doc, update = TRUE, value)
setProperty(doc, node.expression, value, update = inherits(doc, "WordArchive"))
}

\arguments{\item{doc}{the object identifying the document of interest.
   This can be an OOXMLArchive or derived object or a parsed XML document.}\item{node.expression}{xpath expression to identify the property of interest}\item{value}{the new value for the property.}\item{update}{a logical value indicating whether to update the archive
      or just the XML content of the document in memory}>
   \item{date}{a date as a string or POSIXt value}}

\value{These are used for their side effect. They generally return the new value.
     }

\references{}

\seealso{}

  \author{Duncan Temple Lang \email{duncan@wald.ucdavis.edu}
}
 
\examples{
  
}

\keyword{IO}
