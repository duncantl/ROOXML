\name{OOXMLArchive-class}
\Rdversion{1.1}
\docType{class}
\alias{OOXMLArchive-class}
\alias{[[,OOXMLArchive,character,missing-method}
\alias{[[,OOXMLArchive,ANY,missing-method}
\alias{getPart,OOXMLArchive-method}

\title{A generic Office Open XML archive document}
\description{
  This class identifies a Zip archive that is an Office Open
  document such as a Word or Excel file. Different packages provide
  the specific classes for the different types of documents with which
  users interact.
  This class is their super-class and provides general functions methods that
  apply to all Office Open documents.
   
}
\section{Objects from the Class}{
  Objects can be created by calls to \code{\link{createOODoc}}
  or via the constructors from the relevant packages,
  e.g. \code{wordDoc} or \code{excelDoc}.
}
\section{Slots}{
	 \describe{
    \item{\code{.Data}:}{the name of the zip/archive file}
    \item{\code{classes}:}{a vector of class names for this object}
    \item{\code{elements}:}{a character vector giving the components/files of
      the archive when the document was read}
  }
}
\section{Extends}{
Class \code{"\linkS4class{ZipFileArchive}"}, directly.
Class \code{"\linkS4class{character}"}, by class "ZipFileArchive", distance 2.
Class \code{"\linkS4class{ZipArchive}"}, by class "ZipFileArchive", distance 2.
Class \code{"\linkS4class{vector}"}, by class "ZipFileArchive", distance 3.
Class \code{"\linkS4class{data.frameRowLabels}"}, by class "ZipFileArchive", distance 3.
Class \code{"\linkS4class{EnumerationValue}"}, by class "ZipFileArchive", distance 3.
}
\section{Methods}{
  \describe{
    \item{[[}{\code{signature(x = "OOXMLArchive", i = "character", j =
	"missing")}: extract the contents of a component/file in the
      archive by its name.}
    \item{[[}{\code{signature(x = "OOXMLArchive", i = "ANY", j =
	"missing")}: 
             extract the contents of a component/file in the archive by
	     either index or name}
	   \item{getPart}{\code{signature(doc = "OOXMLArchive")}:
	      get a part 
	        }
	 }
}
\references{ ~put references to the literature/web site here ~ }
\author{ Duncan Temple Lang}

\seealso{
  \code{\link[Rcompression]{ZipArchive-class}} in the \code{Rcompression} package.
  \code{WordArchive} and \code{ExcelArchive} in the
  \code{RWordXML} and \code{RExcelXML} packages.

  \code{\link{createOODoc}}
}
\examples{
showClass("OOXMLArchive")
}
\keyword{classes}
