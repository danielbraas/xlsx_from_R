library(tidyverse)
library(openxlsx)

# This script will do the same as the "xlsx_with_xlsx_101.R" and "xlsx_with_XLConnect_101.R"
# scripts but will use the openxlsx package instead of the xlsx and XLConnect packages.


#This paragraph creates a Excel file with three sheets, a data frame, a plot amd a hyperlink 

wb <- createWorkbook(title = 'openxlsx_Test1.xlsx')
addWorksheet(wb, sheetName = 'Iris', tabColour = '#4F81BD', gridLines = F)

hs <- createStyle(fontColour = "#ffff00", fgFill = "#4F80BD",
                  halign = "center", valign = "center", textDecoration = "Bold",
                  border = "TopBottomLeftRight")

freezePane(wb, sheet = 'Iris', firstActiveRow = 3, firstCol = TRUE) ## freeze first row and column
writeDataTable(wb, sheet = 'Iris', x = iris, startRow = 2, startCol = 2, 
               headerStyle = hs, colNames = TRUE, rowNames = F,
               tableStyle = "TableStyleMedium2")

setColWidths(wb, sheet='Iris', widths = 'auto', cols = 1:10)

ggplot(iris, aes(Sepal.Length, Sepal.Width, color=Species))+
  geom_point()

insertPlot(wb, sheet='Iris', width=6, height=4, startRow = 2,
           startCol = length(iris)+2)

addWorksheet(wb, 'Plot')
name <- 'test_plot'
wd <- getwd()

png(paste0(name,'.png'), width=400, height=400)
plot(1, 1, pch=19)
dev.off()

writeData(wb, 'Plot', x='This is a boring plot', borders = 'surrounding')
insertImage(wb, sheet='Plot', file = paste0(name,'.png'), startCol = 2, startRow = 2,
            height = 4, width = 4)


x <- paste0(name, '.png')
class(x) <- 'hyperlink'

addWorksheet(wb, 'hyperlink')
#writeData(wb, 'hyperlink', x = x)
writeFormula(wb, 'hyperlink', x = makeHyperlinkString(sheet = 'testing',
  file = paste0(name, '.png'), text='This is a hyperlink to the file'))

saveWorkbook(wb, 'openxlsx_Test1.xlsx', overwrite = T)
