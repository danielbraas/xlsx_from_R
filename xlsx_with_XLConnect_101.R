library(tidyverse)
library(XLConnect)

# This script will do the same as the "xlsx_with_xlsx_101.R" script but will use 
# the XLConnect package instead of the xlsx package.


#This paragraph creates a Excel file with three sheets, a data frame, a plot amd a hyperlink 
wb <- loadWorkbook('xlconnect_Test1.xlsx', create = T)
sheet <- createSheet(wb, name='Iris')

csHeader <- createCellStyle(wb, name = "heading")
setFillPattern(csHeader, fill = XLC$FILL.SOLID_FOREGROUND)
setFillForegroundColor(csHeader, color = XLC$COLOR.GREY_25_PERCENT)
setBorder(csHeader, side = 'all', type = XLC$"BORDER.THICK", color = XLC$"COLOR.BLACK")

mergeCells(wb, sheet = 'Iris', reference = 'A1:H3')

writeWorksheet(wb, sheet='Iris', header=F, 'This is the Iris dataset')

writeWorksheet(wb, sheet='Iris', header=T, data=iris, startRow = 4, startCol = 1)
setCellStyle(wb, sheet='Iris', row=4, col=1, cellstyle = csHeader)

sheet <- createSheet(wb, 'Plot')
name <- 'test_plot'
wd <- getwd()

png(paste0(name,'.png'), width=400, height=400)
plot(1, 1, pch=19)
dev.off()

writeWorksheet(wb, sheet='Plot',data='This is just a dot in a plot',startRow = 1, startCol = 1, header=F)
setCellStyle(wb, formula = 'Plot!$A$1', cellstyle = csHeader)

createName(wb, name = "plot", formula = "Plot!$B$2", overwrite = TRUE)
addImage(wb, filename=paste0(name,'.png'), name = 'plot', originalSize = T)

sheet <- createSheet(wb, 'hyperlink')

writeWorksheet(wb, sheet='hyperlink', data='This is a hyperlink', header=F)
setHyperlink(wb, formula = 'hyperlink!$A$1', type = XLC$HYPERLINK.FILE, address = paste0(name,'.png'))

saveWorkbook(wb)

clearRange(wb, sheet='Iris', c(2,1,2,1))
