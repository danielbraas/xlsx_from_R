library(xlsx)

#This paragraph creates a Excel file with three sheets, a data frame, a plot amd a hyperlink 
wb <- createWorkbook()
sheet <- createSheet(wb, 'Iris')
addDataFrame(iris, sheet, startRow = 2, startColumn = 2)

sheet <- createSheet(wb, 'Plot')
name <- 'test_plot'
wd <- getwd()

png(paste0(name,'.png'), width=400, height=400)
plot(1, 1, pch=19)
dev.off()

addPicture(paste0(name,'.png'), sheet)

sheet <- createSheet(wb, 'hyperlink')
rows <- createRow(sheet, 1:10)
cells <- createCell(rows, colIndex = 1)
cell <- cells[[1,1]]
setCellValue(cell, paste0(name, '.png'))
addHyperlink(cell, paste(wd,paste0(name,'.png'), sep='/'))

saveWorkbook(wb, 'xlsx_Test1.xlsx')

# This paragraph creates an Excel file with one sheet containing a data frame and a hyperlink next to it

wb <- createWorkbook()
sheet <- createSheet(wb, 'Iris')
addDataFrame(iris, sheet, startRow = 2, startColumn = 2)

rows <- createRow(sheet, rowIndex = 2)
cells <- createCell(rows, colIndex = 8)
cell <- cells[[1,1]]
setCellValue(cell, paste0(name, '.png'))
addHyperlink(cell, paste(wd,paste0(name,'.png'), sep='/'))

saveWorkbook(wb, 'xlsx_Test2.xlsx')

#This paragraph creates a Excel file with three sheets, a data frame, a plot amd a hyperlink 
wb <- createWorkbook()
sheet <- createSheet(wb, 'Iris')
rows <- createRow(sheet, rowIndex = 2)
cells <- createCell(rows, colIndex = 3)
cell <- cells[[1,1]]
red <- CellStyle(wb)+Fill(foregroundColor = 'red')
setCellValue(cell, 'Testing this')
addHyperlink(cell, 'https://systems.crump.ucla.edu', hyperlinkStyle = red)

addDataFrame(iris, sheet, startRow = 2, startColumn = 4, row.names = F)

saveWorkbook(wb,'xlsx_DF_test.xlsx')
