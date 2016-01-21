#' Package defaults
#'
#' This function returns a list of default parameters for this package
lGetPackageDefaults <- function(){
  return(list(sOutFile            = "create_wb_result.xlsx",
              sSheetName          = "sheet1",
              sTableStyle         = "TableStyleMedium17",
              nHeaderTextRotation = 90,
              nStartRow           = 1,
              nStartCol           = 1))
}
