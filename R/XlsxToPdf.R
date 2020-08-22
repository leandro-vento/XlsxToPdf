# Hello, world!
#
# This is an example function named 'hello'
# which prints 'Hello, world!'.
#
# You can learn more about package authoring with RStudio at:
#
#   http://r-pkgs.had.co.nz/
#
# Some useful keyboard shortcuts for package authoring:
#
#   Install Package:           'Ctrl + Shift + B'
#   Check Package:             'Ctrl + Shift + E'
#   Test Package:              'Ctrl + Shift + T'

XlsxToPdf <- function(x, Sheet, Output) {

  ScriptVbs <- ""

  ScriptVbs <- paste0(ScriptVbs, "Dim xlApp\n")
  ScriptVbs <- paste0(ScriptVbs, "Dim xlBook\n")
  ScriptVbs <- paste0(ScriptVbs, 'Set xlApp = CreateObject("Excel.Application")\n')
  ScriptVbs <- paste0(ScriptVbs, "xlApp.Visible = False\n")
  ScriptVbs <- paste0(ScriptVbs, 'Set xlBook = xlApp.Workbooks.Open("', x, '")\n')
  ScriptVbs <- paste0(ScriptVbs, 'xlBook.Sheets("', Sheet, '").ExportAsFixedFormat xlTypePDF, "', Output, '", xlQualityStandard, , , , , False\n')
  ScriptVbs <- paste0(ScriptVbs, "xlApp.Quit\n")
  ScriptVbs <- paste0(ScriptVbs, "Set xlBook = Nothing\n")
  ScriptVbs <- paste0(ScriptVbs, "Set xlApp = Nothing\n")

  write.table(ScriptVbs, file = paste0(getwd(), "\\configura.vbs"), sep = "\n",
              row.names = FALSE, col.names = FALSE, quote = FALSE)


  pathofvbscript <- paste0(getwd(), "\\configura.vbs")
  shell(shQuote(normalizePath(pathofvbscript)), "cscript", flag = "//nologo")

  file.remove(paste0(getwd(), "\\configura.vbs"))

  print("PDF criado com sucesso!")

}
