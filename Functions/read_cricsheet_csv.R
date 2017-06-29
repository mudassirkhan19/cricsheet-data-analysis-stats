readCricsheetCsv <- function(directory = ".", start_no = 0){
    require(data.table)
    all.data.files <- list.files(path = directory, 
                                pattern = "^[^info].*csv$", full.names = TRUE)
    all.info.files <- list.files(path = directory, 
                                 pattern = "^[info].*csv$", full.names = TRUE)
    match_data <- data.table()
    info <- data.table()
    for (i in 1:length(all.data.files)) {
        new_data <- fread(all.data.files[i],na.strings = c("NA",""))
        new_data[, match_no := start_no + i]
        match_data <- rbind(match_data, new_data,fill=TRUE)
    }
    for (i in 1:length(all.info.files)) {
        new_info <- fread(all.info.files[i],na.strings = c("NA",""))
        new_info[, match_no := start_no + i]
        info <- rbind(info, new_info,fill=TRUE)
    }
    
    return(list(match_data,info))
}