modify_excel_sheet <- function(sheetName, wbPath, data){
    require(openxlsx)
    if(!file.exists(wbPath)){
        write.xlsx(data, wbPath, sheetName = sheetName)
    }
    else{
        wb <- loadWorkbook(wbPath)
        if(sheetName %in% getSheetNames(wbPath)){
            removeWorksheet(wb, sheetName)
        }
        addWorksheet(wb, sheetName)
        writeData(wb, sheetName, data)
        saveWorkbook(wb,wbPath,overwrite = T)
    }
}