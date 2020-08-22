#' Transforma um arquivo xlsx em pdf.
#'
#' @param x É o nome e caminho onde está localizado o arquivo xlsx.
#' @param Sheet É o nome da guia no arquivo xlsx que deseja salvar como pdf.
#' @param Output É o nome e caminho onde será salvo o arquivo pdf
#' @examples
#' XlsxToPdf("C:/MeuArquivo.xlsx", "Plan1", "C:/MeuArquivo.pdf")

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