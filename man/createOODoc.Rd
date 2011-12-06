
\name{createOODoc}
\alias{createOODoc}

\title{Create object representing an Office Open document}

\description{This is a general function that can be used by other packages
and users to create an object that represents an Office Open
document such as a Word or Excel document.
The function allows the caller to specify the class name of the target object
and the general OO archive values are filled in.
     }

\usage{
createOODoc(f, create = FALSE, class = "OOXMLArchive", emptyDoc = NA, extension = character())
}

\arguments{\item{f}{the name of the Office Open file}\item{create}{a logical value indicating whether we should create the file if it does not exist (by copying a template)}\item{class}{the name of the class of which we want an instance.  This allows the caller to use this function
to create instances of, e.g., the WordArchive or ExcelArchive classes. This should be the name of a class that extends
OOXMLArchive.}\item{emptyDoc}{the name of a template document of this type if the file given via f does not exist and
create is \code{TRUE}.}\item{extension}{a string giving the extension to check for if only the base name of the document is given}}

\value{An object of class given by  class.
This has the name of the file and the vector of names 
of the components within the archive.
     }

\references{}

\seealso{\code{\link[RWordXML]{wordDoc}} and \code{\link[RExcelXML]{excelDoc}}}

  \author{Duncan Temple Lang \email{duncan@wald.ucdavis.edu}
}
 
\examples{
  
}

\keyword{}
