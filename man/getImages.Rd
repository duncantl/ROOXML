
\name{getImages}
\alias{getImages}
\alias{getImages,XMLInternalDocument-method}

\title{Identify the components in an Office Open archive that are images within the document}

\description{This function examines the Office Open archive and 
     queries the names of the image files that it contains and are referenced
     within the "document".
     This is  a generic function and different packages provide methods
     for different types of Office Open archives. That is RExcelXML
      and RWordXML provide methods.
      These methods do more than just look in the zip file for files with standard image 
    extensions. Instead, they look at the document and its relationships file and finds all
    the files in the archive corresponding to the image type.

     }

\usage{
getImages(doc, ...)
}

\arguments{\item{doc}{the Office Open document archive.
This is an instance of a class derived from OOXMLArchive,
e.g. WordArchive or ExcelArchive}\item{\dots}{additional parameters for methods}}

\value{a character vector giving the names of the files/components within the archive corresponding
to the images in the document. The names of the vector give the relational identifiers
    used in the content of the document to refer to the image files. See the relationships
part of the document.
     }

\references{}

\seealso{}

  \author{Duncan Temple Lang \email{duncan@wald.ucdavis.edu}
}
 
\examples{
   #  A Word Document
 #  This illustrates how to access the images in a Word document. This belongs in RWordXML!
if(require(RWordXML)) {
   if(file.exists(system.file("SampleDocs", "Images2.docx", package = "RWordXML"))) {
     doc = wordDoc(system.file("SampleDocs", "Images2.docx", package = "RWordXML"))
     getImages(doc)
   }

  doc = wordDoc(system.file("SampleDocs", "pics.docx", package = "ROOXML"))
  getImages(doc)
}
         #  This version illustrates extracting the names of the image components from
 #  an Excel document.
if(require(RExcelXML)) {
   e = excelDoc(system.file("SampleDocs", "Images.xlsx", package = "ROOXML"))
   getImages(e)
}

}

\keyword{documentation}

\keyword{IO}
