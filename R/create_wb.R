#' Create an excel workbook
#'
#' \code{create_excel_wb} creates an excel workbook from an R dataframe
#'
#' The only input that must be specified is the dataframe to be converted or
#' a csv file that can be read into a dataframe. From that an excel workbook
#' is created using some standard format.
#'
#' @param pdInput                input data frame
#' @param psInFile               name of the input file
#' @param psOutFile              name of the output file
#' @param psSheetName            name of the sheet to be used
#' @param psTableStyle           pre-defined table style to be used
#' @param pnHeaderTextRotation   rotation angle for header text
#' @export
#' @examples
#' create_excel_wb(pdInput = iris)
#'
create_excel_wb <- function(pdInput, psInFile,
                            psOutFile = NULL,
                            psSheetName = NULL,
                            psTableStyle = NULL,
                            pnHeaderTextRotation = NULL,
                            pnStartRow = NULL,
                            pnStartCol = NULL){
  # in case no input dataframe pdInput was specified, try to read it from psInFile
  if (is.null(pdInput)){
    stopifnot(file.exists(psInFile))
    dfInput <- read.csv2(file = psInFile)
  } else {
    dfInput <- pdInput
  }
  # get the defaults for all other parameters
  lDefaults <- lGetPackageDefaults()
  # replace parameter with defaults, if they are not specified
  sOutFile <- ifelse (is.null(psOutFile),  lDefaults$sOutFile, psOutFile)
  sSheetName <- ifelse (is.null(psSheetName), lDefaults$sSheetName, psSheetName)
  sTableStyle <- ifelse (is.null(psTableStyle), lDefaults$sTableStyle, psTableStyle)
  nHeaderTextRotation <- ifelse (is.null(pnHeaderTextRotation), lDefaults$nHeaderTextRotation, pnHeaderTextRotation)
  nStartRow <- ifelse(is.null(pnStartRow), lDefaults$nStartRow, pnStartRow)
  nStartCol <- ifelse(is.null(pnStartCol), lDefaults$nStartCol, pnStartCol)
  # create workbook and save it to result file
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, sSheetName)
  openxlsx::writeDataTable(wb, 1, dfInput, startRow = nStartRow, startCol = nStartCol,
                           tableStyle = sTableStyle,
                           headerStyle = openxlsx::createStyle(textRotation = nHeaderTextRotation))
  openxlsx::saveWorkbook(wb, sOutFile, overwrite = TRUE)
  invisible()
}
