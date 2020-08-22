#' Transforma um arquivo xlsx em pdf.
#'
#' @param x Nome e caminho onde está localizado o arquivo xlsx.
#' @param Sheet Nome da guia no arquivo xlsx que deseja salvar como pdf.
#' @param Output Nome e caminho onde será salvo o arquivo pdf
#' @examples
#' XlsxToPdf("C:/MeuArquivo.xlsx", "Plan1", "C:/MeuArquivo.pdf")
#' @export
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

  utils::write.table(ScriptVbs, file = paste0("configura.vbs"), sep = "\n",
              row.names = FALSE, col.names = FALSE, quote = FALSE)


  pathofvbscript <- paste0("configura.vbs")
  shell(shQuote(normalizePath(pathofvbscript)), "cscript", flag = "//nologo")

  file.remove(paste0("configura.vbs"))

  print("PDF criado com sucesso!")

}
