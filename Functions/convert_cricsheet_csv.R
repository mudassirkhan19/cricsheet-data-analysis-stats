aggr_fielder <- function(x) {
    paste0(x, collapse="/")
}

convertCricsheetData <- function(source = ".",destination = ""){
    require(yaml)
    require(reshape2)
    require(data.table)
    require(stringr)
    all.files <- list.files(path = source,
                            pattern = ".yaml$",
                            full.names = TRUE)

    for (i in 1:length(all.files)) {
        data <- yaml.load_file(all.files[i])
        x <- melt(data)
        y <- data.table(x)
        # Fix for winner = 1 bug
        y[L2 == "outcome" & L3 == "winner", value := data$info$outcome$winner]
        # extracting file no from name
        fileName <- gsub("\\.yaml$","",all.files[i])
        matches <- regmatches(fileName, gregexpr("[[:digit:]]+$", fileName))
        fileName <- unlist(matches)
        # Fixing bug for 10th ball of an over
        while (nrow(y[(grepl("\\.9$|\\.10$",shift(L6))) & grepl("\\.1$",L6) & (str_split(L6, "\\.", simplify = T)[,1] == str_split(shift(L6), "\\.", simplify = T)[,1]), ]) > 0) {
            y[ grepl("\\.9$|\\.10$",shift(L6)) & grepl("\\.1$",L6) & (str_split(L6, "\\.", simplify = T)[,1] == str_split(shift(L6), "\\.", simplify = T)[,1]),
               L6 := paste0(L6,"0")]
        }
        
        # meta = y[y$L1 == 'meta',]
        # meta = meta[, colSums(is.na(meta)) != nrow(meta), with=FALSE]
        # data_meta = reshape(meta,direction = 'wide',timevar = 'L2',idvar = 'L1')

        info = y[y$L1 == 'info',]
        info = info[, colSums(is.na(info)) != nrow(info), with=FALSE]
        info[, L1 := NULL]
        
        data_innings = y[(y$L1 == 'innings') & (y$L4 == 'deliveries'),]
        data_innings[, new := paste(data_innings$L7,data_innings$L8,sep="_")]
        
        data_innings$over <- sapply(strsplit(data_innings$L6, "[.]"), `[`, 1)
        data_innings$over <- as.numeric(data_innings$over) + 1
        data_innings$ball <- sapply(strsplit(data_innings$L6, "[.]"), `[`, 2)
        
        data_innings [, c("L7","L8","L4","L1","L5","L6") := NULL]
        # Fix for two or more fielders in a runout
        data_innings = dcast(data_innings, L2+L3+over+ball ~ new, fun.aggregate = aggr_fielder,fill = "",value.var = "value")
        fwrite(data_innings,paste0(destination,paste(c(info[info$L2 == "dates",]$value,info[info$L2 == "teams",]$value,fileName), collapse = "-"),".csv"),row.names = F)
        fwrite(info,paste0(destination,paste(c("info",info[info$L2 == "dates",]$value,info[info$L2 == "teams",]$value,fileName), collapse = "-"),".csv"),row.names = F)
    }
    
}
