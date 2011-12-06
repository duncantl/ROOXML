if(require(RWordXML)) {
if(!file.exists("foo.R"))
   cat("f =\nfunction() {\n 1\n}\n", file = "foo.R")

library(ROOXML)


#ar = wordDoc("foo.docx")
ar = createOODoc("foo.docx")
names(ar)

ar[["x/bob"]] = "foo.R"
ar[["x/bob"]]

ar = wordDoc("bar.docx")
ar[c("x/bob", "jane")] = list("foo.R", I("inlined\ncontent"))
names(ar)
}
