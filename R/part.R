if(FALSE) {
xx = xpathApply(xmlParse(ar[["[Content_Types].xml"]]), "//w:Override", function(x) sapply(c("PartName", "ContentType"), function(id) xmlGetAttr(x, id)), namespaces = "w")
dput(structure(sapply(xx, `[`, 2), names = sapply(xx, `[`, 1)))
}

WordprocessingMLParts =
structure(c("application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", 
"application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml", 
"application/vnd.openxmlformats-package.core-properties+xml", 
"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml", 
"application/vnd.openxmlformats-officedocument.theme+xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", 
"application/vnd.openxmlformats-officedocument.extended-properties+xml", 
"application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
"relationships"            
), .Names = structure(c("/word/styles.xml", "/word/webSettings.xml", 
"/docProps/core.xml", "/word/settings.xml", "/word/theme/theme1.xml", 
"/word/document.xml", "/docProps/app.xml", "/word/fontTable.xml", "/word/_rels/document.xml.rels"
)))     
          

setGeneric("getOOXMLPartsTable", function(obj) standardGeneric("getOOXMLPartsTable"))

    

setGeneric("getPart",
            function(doc, part, default = NA, stripSlash = TRUE, parts = computeParts(doc)) # getOOXMLPartsTable(doc))
             standardGeneric("getPart"))

#setMethod("getPart", "OOXMLArchive",
#function(doc, part, default = NA, stripSlash = TRUE, parts = WordprocessingMLParts)
#{
#  getPart(doc, part, default, stripSlash, WordprocessingMLParts)
#})
     
setMethod("getPart", "OOXMLArchive",
  #
  # Lookup the name of the entry in the archive corresponding to a
  # conceptual part, e.g. document.
  # Can index with the full name of the part or the name of the
  # usual file.
  #
  # We now ignore the static table, and compute the part mapping of type -> file
  # via computeParts. Then we look up [Content_Types] again. Don't do this.       
  #
  #
function(doc, part, default = NA, stripSlash = FALSE, parts = computeParts(doc)) #getOOXMLPartsTable(doc)) # WordprocessingMLParts
{
   orig.part = part
      # Find the name of the thing we want.
   i = pmatch(part, parts)
   if(is.na(i))
      i = pmatch(part, names(parts))

   if(is.na(i)) {
       # match by regular expression
     i = grep(part, parts, fixed = inherits(part, "AsIs"))
     if(length(i) > 1)
        stop("ambiguous part ", part, " matches ", paste(parts[i], collapse = ", "))

     if(length(i) == 0)
       i = grep(part, names(parts), fixed = inherits(part, "AsIs"))

     if(length(i) > 1)
        stop("ambiguous part ", part, " matches ", paste(names(parts)[i], collapse = ", "))     
   }
   

   if(length(i)) {
     part = parts[ i ]     
     tmp = names(part)
     if(stripSlash)
       tmp = gsub("^/", "", tmp)

      return( tmp )
   }
   
    # Look in the [Content_Types].xml file.
    # We can ignore this now since computeParts() has already done that.
if(FALSE) {
   x = doc[["[Content_Types].xml"]]   
   n = getNodeSet(x, paste("//x:Override[@ContentType =", sQuote(part), "]"), "x")
   if(length(n))  {
       tmp = xmlGetAttr(n[[1]], "PartName")
       if(stripSlash)
         gsub("^/", "", tmp)
       else
         tmp
   }
}

       if(orig.part %in% parts)
          names(parts)[match(orig.part, parts)]
       else
         default   
})

setGeneric("getDocument",
            function(doc, ...)
               standardGeneric("getDocument"))




computeParts =
function(doc)
{
   typeDoc = contentTypes(doc)
   nodes = getNodeSet(typeDoc, "/*/x:Override", "x")
   structure(sapply(nodes, xmlGetAttr, "ContentType"),
             names = sapply(nodes, xmlGetAttr, "PartName"))
}



setMethod("[[", c("OOXMLArchive", j = "missing"),
  # ... allows password and mode.
  # i can be a character or a number
  # Need to implement as number.
function(x, i, j, ..., password = character(), mode = "", nullOk = TRUE)
{ 
  getZipFileEntry(x, i, password, mode, nullOk = nullOk)
})
