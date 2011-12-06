
sQuote = function(x)
   sprintf("'%s'", x)

#library(Rcompression)
setClass("OOXMLArchive", contains = "ZipFileArchive") #n, prototype = prototype(classes = extends("OOXMLArchive")))
#setClass("WordArchive", contains = "OOXMLArchive") #, prototype = list(classes = extends("WordArchive")))
#setClass("ExcelArchive", contains = "OOXMLArchive")# , prototype = list(classes = extends("ExcelArchive")))



#"[[.OOXMLArchive" =
setMethod("[[", c("OOXMLArchive", "character", j = "missing"),
function(x, i, j, ..., asXML = TRUE)
{
   ans = callNextMethod() #  NextMethod("[[")
     # perhaps we want to check the types of the document
     # via the [Content_Types].xml file.
   if(asXML && is.character(ans) && (length(grep("\\.xml$", i)) > 0 ||
       (length(type <- getEntryType(i, x, partialMatch = TRUE)) > 0 && grepl("\\+xml$", type))))
     ans = xmlParse(ans, asText = TRUE)

   if(inherits(ans, "XMLInternalDocument"))
     docName(ans) = paste(x, i, sep = "::")

   ans
})



createOODoc =
function(f, create = FALSE, class = "OOXMLArchive", emptyDoc = NA, extension = character() )  
{
  if(!file.exists(f) && file.exists(paste(f, extension, sep = ".")))
    f = paste(f, extension, sep = ".")
  
  if(!file.exists(f) && create && !is.na(emptyDoc)) {
      #XXX Make certain to remove the identifying meta-data/information
      # and update it with info for this user.
    file.copy(emptyDoc, f)
  }
    
  ans = zipArchive(f, class = class)
  new(class, ans@.Data, elements = ans@elements, classes = c(class, if(length(ans@classes)) ans@classes else extends(class)[-1]),
         readTime = ans@readTime)
}

creator = 
function(doc)
  property(doc, "creator")

subject = 
function(doc)
  property(doc, "subject")

description = 
function(doc)
  property(doc, "description")

keywords = 
function(doc, sep = "[,;]")
{
  ans = property(doc, "keywords")
  if(is.null(ans) || is.na(sep))
    return(character())
  
  unlist(strsplit(ans, sep))
}

Title = title = 
function(doc)
  property(doc, "title")

revision = 
function(doc)
  property(doc, "revision")

created = 
function(doc)
  strptime(property(doc, "created"), "%Y-%m-%dT%H:%M:%SZ")

modified = 
function(doc)
  strptime(property(doc, "modified"), "%Y-%m-%dT%H:%M:%SZ")

category = 
function(doc)
  property(doc, "category")

property = 
function(doc, el,  tt = doc[["docProps/core.xml"]])
{
  properties(doc, tt = tt)[[el]]
}

properties = 
function(doc, ..., tt = doc[[if(custom) "docProps/custom.xml" else "docProps/core.xml"]], 
          .els = as.character(unlist(list(...))), custom = FALSE) # ? What is .els!
{
  if(is.null(tt))
     return(list())
  
  ans = if(custom) {
            customProperties(tt)
        } else
            xmlToList(tt)

  if(length(.els))
     ans[.els]
  else
     ans
}

customProperties =
function(customDoc)
{
  if(is.null(customDoc))
     return(list())
  
  structure(xmlApply(xmlRoot(customDoc), function(x) valueOfNode(x[[1]])),
             names = xmlSApply(xmlRoot(customDoc), xmlGetAttr, "name"))
}

valueOfNode =
function(n)
{
   # See http://schemas.liquid-technologies.com/OfficeOpenXML/2006/vector4.html
   # Perhaps use XMLSchema to convert this.
  val = xmlValue(n)
  switch(xmlName(n),
          "i4" = as.integer(val),
          "i2" = as.integer(val),
          "bool" = as.logical(val),
          "r8" = as.numeric(val),
          "r4" = as.numeric(val),
          "date" = strptime(val, "%Y-%m-%d"),
          "dateTime" = strptime(val,  "%Y-%m-%dT%H:%M:%SZ"),
          "filetime" = strptime(val,  "%Y-%m-%dT%H:%M:%SZ"),         
         val)         
}

commentNodes =
function(doc)
{
  comments(doc, function(x) x)
}

comments = 
function(doc, op = xmlValue)
{
  tt =  doc[[getPart(doc, I("comments+xml"))]]
  if(!is.null(op))
    xmlSApply(xmlRoot(tt), op)
  else
    xmlChildren(xmlRoot(tt))
}

commentTable = 
function(doc)
{
  tt = doc[[getPart(doc, I("comments+xml"))]]
  root = xmlRoot(tt)
  vars = c("id", "author", "date")
  ans = as.data.frame(lapply(vars,
                             function(at) 
                              xmlSApply(root, function(node) 
                                                xmlGetAttr(node, at, ""))))
  names(ans) = vars
  ans$date = as.POSIXct(strptime(ans$date, "%Y-%m-%dT%H:%M:%SZ"))
  ans$value = xmlSApply(root, xmlValue)
  ans
}


commentTable2 =
  #
  # From Gabe.
  #
  #
  # Unlike above, this function grabs information about the styles within the
  # comments.
  #
  # NOTE 1: The resulting data.frame has one row for each paragraph in the
  #         comments file, NOT one row per id (comment) as in the function
  #         above.
  # NOTE 2: This function only detects paragraph level styles. Inline styles
  #         will be ignored.
function(doc)
{
  i = getPart(doc, I("comments+xml"))
  if(is.na(i))
      return(NULL)
  
  tt = doc[[i]]
  root = xmlRoot(tt)
  paras = getNodeSet(root, ".//w:comment/w:p")
  vars = c("id", "author", "date")
  ans = as.data.frame(t(sapply(paras, function(node)
          lapply(vars, function(var) xmlGetAttr(xmlParent(node),var, "")))))
  names(ans) = vars
  ans$style = xpathApply(root, "//w:comment/w:p/w:pPr/w:pStyle/@w:val")
  ans$value = sapply(paras, xmlValue)
  ans
}


setGeneric("hyperlinks",
            function(doc, comments = FALSE, ...)
             standardGeneric('hyperlinks'))


# Get the name of the element/file in the archive, not of the archive
# itself.
setGeneric("getDocElementName", function(doc, addSlash = FALSE, ...)
               standardGeneric('getDocElementName'))

setMethod("getDocElementName", "XMLInternalDocument",
          function(doc, addSlash = FALSE, ...) {
            els = docName(doc)
            els = strsplit(els, "::")[[1]]
            ans = if(length(els) > 1)
                    els[2]
                  else
                    els[1]
            if(addSlash)
              sprintf("/%s", ans)
            else
              ans
          })


# Get the relationships document for a given document,
# be it a zipArchive, OOXMLArchive or an element within that
setGeneric("getRelationshipsDoc",
            function(doc, ...) {
              standardGeneric("getRelationshipsDoc")
            })

setMethod("getRelationshipsDoc", "XMLInternalDocument",
            function(doc, ...) {
              tmp = getDocElementName(doc)
              sprintf("%s/_rels/%s.rels", dirname(tmp), basename(tmp))
            })

setMethod("getRelationshipsDoc", "OOXMLArchive",
            function(doc, ...) {
               "_rels/.rels"
            })

#???? is rels the _rels/.rels or from the document word/_rels/document.xml.rels
#  but that is Word specific.
"getRelationships" =
function(doc, rels = doc[["_rels/.rels"]]) # getPart(doc, "relationships"))
{
    # Now lookup the target URLs, i.e. the hrefs.
    # There may be more in that document as comments might 
    # have hyperlinks. So we need the subset in the xml.rels 
    # document that match the identifiers for our w:hyperlinks.
#  rels = doc[[ rels ]] 
  ids = xmlSApply(xmlRoot(rels), xmlGetAttr, "Id")
  structure(data.frame(id = ids, target =  xmlSApply(xmlRoot(rels), xmlGetAttr, "Target"),
                       type =  xmlSApply(xmlRoot(rels), xmlGetAttr, "Type"),
                       stringsAsFactors = FALSE),
            class = c("OORelationshipsTable", "data.frame"))
}

resolveRelationships =
function(ids, doc, relTable = getRelationships(doc), relativeTo = character())
{
  val = relTable$target[ match(ids, relTable$id) ]
  if(length(relativeTo) && !is.na(relativeTo))
     val = paste(relativeTo, val, sep = "/")
  structure(val,  names = ids)
}



setProperty =
function(doc, node.expression, value, update = inherits(doc, "WordArchive"))
{
  part = NA
  if(inherits("WordArchive")) {
    part = getPart(doc, "docProps/core.xml")
    xml = doc[[part]]
  } else
     xml = doc

  mod = getNodeSet(xml, node.expression)
 
  if(length(mod) == 0)
    newXMLNode(gsub(".*/", "", node.expression), value, parent = xmlRoot(xml))
  else
    xmlValue(mod[[1]]) = value

  if(update && !is.na(part))
    doc[[part]] = xml
  else
    xml
}

'creator<-' =
function(doc, update = TRUE, value)
{
  setProperty(doc, "//dc:creator", value, update)
}

setUpdated =
  #
  #  Update the date for the modified field and increment the revision.
  #
function(doc, date = Sys.time(), update = TRUE)
{
  val = format(date, "%Y-%m-%dT%H:%M:%SZ")
  doc = setProperty(doc, "//dcterms:modified", val, FALSE)
  setProperty(doc, "//cp:revision",   as.integer(xmlRoot(doc)[["revision"]]) + 1L, update)  
}


setGeneric("getImages", function(doc, ...) standardGeneric("getImages"))


#setMethod("getImages", "character", 
#           function(doc, ...)
#               getImages(zipArchive(doc), ...))



setMethod("getImages", "XMLInternalDocument",
           function(doc, ...) {
             tt = doc
             imgs = xmlSApply(xmlRoot(tt), xmlGetAttr, "Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
             structure(sapply(xmlRoot(tt)[imgs], xmlGetAttr, "Target"),
                       names = sapply(xmlRoot(tt)[imgs], xmlGetAttr, "Id"))
          })



contentTypes =
  #
  #
  #
function(ar, asXML = TRUE)
{
  ar[["[Content_Types].xml", asXML = asXML]]
}


getEntryType =
function(filename, archive, types = computeParts(archive), partialMatch = FALSE) # contentTypes(archive))
{
#  partNodes = getNodeSet(types, "//x:Override", namespaces = "x")
#  part = sapply(partNodes, xmlGetAttr, "PartName")
  i = match(filename, names(types))
  if(is.na(i)) i = match(paste("/", filename, sep = ""), names(types))

  if(!is.na(i)) return(types[i]) # return(xmlGetAttr(partNodes[[i]], "ContentType"))

  if(partialMatch) {
    i = grep(filename, names(types), fixed = TRUE)
    if(length(i)) {
       if(length(i) > 1) {
         warning("matched more than one archive elements. Using longest match")
         m = regexpr(filename, names(types), fixed = TRUE)
         i = which.max(attr(m, "match.length"))
       }
       return(types[i])
    }
  }

  getEntryTypeByExtension(filename, archive)
}

getEntryTypeByExtension =
function(filename, archive, types = contentTypes(archive), partialMatch = FALSE) 
{  
    # Get the extension.
  ext = gsub(".*\\.([a-zA-Z]+)$", "\\1", filename)
  extNodes = getNodeSet(types, "//x:Default", namespaces = "x")
  i = match(ext, sapply(extNodes, xmlGetAttr, "Extension"))
  if(!is.na(i))
     xmlGetAttr(extNodes[[i]], "ContentType")
  else
    character()
}


if(!isGeneric("author"))
  setGeneric("author", function (doc, ...)  standardGeneric("author"))

setMethod("author", "OOXMLArchive",
           function (doc, ...)  {
              creator(doc)
           })


addExtensionTypes =
  #
  #  Add the <Default Extension="" ContentType="" />
  #  nodes to [Content_Types].xml
  #
function(archive, exts, update = TRUE)  
{
   ctp = archive[["[Content_Types].xml"]]
   files = list()
   if(any(mapply(addDefaultContentType, names(exts), exts, MoreArgs = list(xmlRoot(ctp)))))
      files[["[Content_Types].xml"]] = ctp
   
   if(update)
      updateArchiveFiles(archive, list("[Content_Types].xml" = I(saveXML(ctp))))

   invisible(ctp)
}

addDefaultContentType =
function(ext, type, node) 
{
   if(length(getNodeSet(node, paste(".//x:Default[@Extension = ", sQuote(ext), "]"), "x")) == 0) {
      newXMLNode("Default", attrs = c(Extension = ext, ContentType = type), parent = node)
      TRUE
    } else
      FALSE
}



setGeneric("addToDoc",
function(archive, node, update = TRUE, ...)
           standardGeneric("addToDoc"))



relativeTo =
function(filename, dir, sep = .Platform$file.sep)
{
  els = strsplit(filename, sep)[[1]]
  if(els[1] == dir)
    els = els[-1]
  paste(els, collapse = sep)
}



####################################################################
#
# Insert an image into the archive.

# addImage is now in RWordXML. We will try to bring back the general
# parts into ROOXML.
if(FALSE)
addImage =
  #
  #  This updates the document archive by adding the specified
  #  image file to the archive, ensuring that the Defaults are in 
  #  the [Content_Types].xml file
  #
  #  if node is an XMLInternalDocument, then add it to the end.
  #
  #
function(archive, filename, node = paragraph(align = "center"), ..., after = TRUE, sibling = NA)
{
   newName = paste("word/media", basename(filename), sep = .Platform$file.sep)

   files = structure(list(filename), names = newName)

     # Add a new relationship to the document.xml.rels that identifies the new image
     # file and has an associated id so we can refer to it in the drawing node.
   rels = archive[["word/_rels/document.xml.rels"]]
     # XXX should test this is "available/free", i.e. not being used already.
   rid = paste("rId", xmlSize(xmlRoot(rels)) + 1, sep = "")
   addChildren(xmlRoot(rels),
                newXMLNode("Relationship",
                            attrs = c(Id= rid,
                                      Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                                      Target = relativeTo(newName, "word"))))
   files[["word/_rels/document.xml.rels"]] = rels

        #XXX Need to add to the \[Content_Types].xml file and put in the
        #  entries <Default Extension="pdf" ContentType="application/pdf"/>
        # Could be smarter and just add the type of the image we are dealing with.
        # Error if we add one that is already there.
   addExtensionTypes(archive, c(pdf = "application/pdf", jpeg = "image/jpeg", png = "image/png", jpg = "image/jpeg"))

     # Add a new node
   if(!is.null(node)) {
      if(inherits(node, "XMLInternalDocument")) 
         node = xmlRoot(node)

      imgNode = createImageNode(rid, filename, ...)

        # If sibling is a node, add imgNode as sibling of that
        # if sibling is a logical, then add imgNode as a sibling of
        # node.  Otherwise add imgNode as a child of node.
      if(inherits(sibling, "XMLInternalNode")) {
         addChildren(node, imgNode)
         addSibling(sibling, node) #XXX needs to be immediately after this sibling node, not at end.
      } else if(is.logical(sibling) && !is.na(sibling) && sibling)
        addSibling(node, imgNode)
      else
        addChildren(node, imgNode)
   }
   
   updateArchiveFiles(archive, as(node, "XMLInternalDocument"), .files = files)
   if(!is.null(node))
      return(imgNode)
   else
      rid
}
