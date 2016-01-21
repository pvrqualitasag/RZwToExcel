## ----InstallOpenxlsx, eval=FALSE-----------------------------------------
#  install.packages(pkgs = "openxlsx")

## ----BrowseVignettes, eval=FALSE-----------------------------------------
#  browseVignettes(package = "openxlsx")

## ----WriteIrisToXlsx-----------------------------------------------------
library(openxlsx)
write.xlsx(iris, file = "writeIrisToXLSX1.xlsx")

## ----HeaderStyle---------------------------------------------------------
hs <- createStyle(fontColour = "#ffffff", fgFill = "red",
                  halign = "center", valign = "center", textDecoration = "Bold",
                  border = "TopBottomLeftRight", textRotation = 90)
write.xlsx(iris, file = "writeXLSX4.xlsx", borders = "rows", headerStyle = hs)

## ----FormattingWriteDataWriteDataTable-----------------------------------
## data.frame to write
df <- data.frame("Date"       = Sys.Date()-0:4,
                 "Logical"    = c(TRUE, FALSE, TRUE, TRUE, FALSE),
                 "Currency"   = paste("$",-2:2),
                 "Accounting" = -2:2,
                 "hLink"      = "http://cran.r-project.org/",
                 "Percentage" = seq(-1, 1, length.out=5),
                 "TinyNumber" = runif(5) / 1E9, stringsAsFactors = FALSE)
class(df$Currency) <- "currency"
class(df$Accounting) <- "accounting"
class(df$hLink) <- "hyperlink"
class(df$Percentage) <- "percentage"
class(df$TinyNumber) <- "scientific"
## Formatting can be applied simply through the write functions
## global options can be set to further simplify things
options("openxlsx.borderStyle" = "thin")
options("openxlsx.borderColour" = "#4F81BD")
## create a workbook and add a worksheet
wb <- createWorkbook()
addWorksheet(wb, "writeData auto-formatting")
writeData(wb, 1, df, startRow = 2, startCol = 2)
writeData(wb, 1, df, startRow = 9, startCol = 2, borders = "surrounding")
writeData(wb, 1, df, startRow = 16, startCol = 2, borders = "rows")
writeData(wb, 1, df, startRow = 23, startCol = 2, borders ="columns")
writeData(wb, 1, df, startRow = 30, startCol = 2, borders ="all")
## headerStyles
hs1 <- createStyle(fgFill = "#4F81BD", halign = "CENTER", textDecoration = "Bold",
border = "Bottom", fontColour = "white")
writeData(wb, 1, df, startRow = 16, startCol = 10, headerStyle = hs1,
borders = "rows", borderStyle = "medium")
## to change the display text for a hyperlink column just write over those cells
writeData(wb, sheet = 1, x = paste("Hyperlink", 1:5), startRow = 17, startCol = 14)
1
## writing as an Excel Table
addWorksheet(wb, "writeDataTable")
writeDataTable(wb, 2, df, startRow = 2, startCol = 2)
writeDataTable(wb, 2, df, startRow = 9, startCol = 2, tableStyle = "TableStyleLight9")
writeDataTable(wb, 2, df, startRow = 16, startCol = 2, tableStyle = "TableStyleLight2")
writeDataTable(wb, 2, df, startRow = 23, startCol = 2, tableStyle = "TableStyleMedium21")
# openXL(wb)

## ----OurOwnStyle---------------------------------------------------------
wb <- createWorkbook()
addWorksheet(wb, "IrisOurOwnStyle")
writeDataTable(wb, 1, iris, startRow = 2, startCol = 2, 
               tableStyle = "TableStyleMedium17", 
               headerStyle = createStyle(textRotation = 90))

# openXL(wb)

## ----LargTestData--------------------------------------------------------
devtools::load_all()
dfZw <- read.csv2(file = system.file(file.path("inst","extdata","testdata","csv","NewOrder.csv"), 
                                     package = "RZwToExcel"))
wb <- createWorkbook()
addWorksheet(wb, "Zuchtwerte")
writeDataTable(wb, 1, dfZw, startRow = 1, startCol = 1, 
               tableStyle = "TableStyleMedium17", 
               headerStyle = createStyle(textRotation = 90))
saveWorkbook(wb, "ZwVMS.xlsx", overwrite = TRUE)

