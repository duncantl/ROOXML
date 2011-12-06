
\name{commentTable}
\alias{commentTable}
\alias{commentTable2}
\alias{commentNodes}
\alias{comments}

\title{Access the comments in the document}

\description{These functions allow the caller to retrieve information about the comments in an Office Open document.
We can get just the content of the comments, or the contents along with meta-information
such as the date and author, or the XML nodes representing the comments.
The last of these is useful if we want to programmatically modify the comments.

     \code{commentTable2}
    returns the content of the comments and also the styles associated with those comments.
  This information can be used to interpret the comments, e.g. as R code or R plot or R options.
  This functionality may be merged back into 
     \code{commentTable}
    with a parameter controlling whether the styles are returned.
     }

\usage{
commentTable(doc)
commentTable2(doc)
commentNodes(doc)
comments(doc, op = xmlValue)
}

\arguments{\item{doc}{the Office Open document}\item{op}{ a function that is used to process each of the comment nodes, e.g. xmlValue}}

\value{
     \code{comments}
    returns a character vector giving the contents of the comments.
    
     \code{commentTable}
    returns a data frame with as many rows as there are comments.
The data frame has 4 variables: 
 
 \item{id}{the identifier for the author}
 \item{author}{the name of the author}
 \item{date}{the date the comment was changed and of class POSIXct.}
 \item{value}{the text content of the comment}
     \code{commentNodes}
    returns a list of XMLInternalElementNode objects.
     }

\references{}

\seealso{\code{\link[RWordXML]{wordDoc}}}

  \author{  \email{}
}
 
\examples{
  
     file = system.file("SampleDocs", "WordEg.docx", package = "ROOXML")
if(require(RWordXML)) {
     doc = wordDoc(file)
     comments(doc)

     commentTable(doc)

     commentNodes(doc)
}
        
}

\keyword{documentation}
