cleanScrapedData <- function(){
    require(data.table)
    source('~/RFiles/Projects/ODI/Functions/read_cricsheet_csv.R')
    
    odi_list <- readCricsheetCsv("./match_data_scraped/")
    match_data <- odi_list[[1]]
    info <- odi_list[[2]]
    
    changecols <- setdiff(names(match_data),c("L3","batsman_NA","bowler_NA","non_striker_NA","wicket_fielders","wicket_kind","wicket_player_out"))
    match_data[, (changecols) := lapply(.SD, as.integer), .SDcols = changecols]
    match_data <- match_data[order(match_no, L2, over, ball)]
    
    match_data[, ball := cumsum(!is.na(ball)), by = .(match_no,L2,over)]
    match_data[, over := over + 1]
    source('~/RFiles/Projects/ODI/Functions/matchInfoReshape.R')
    info <- matchInfoReshape(info)
    
    match_played_data <- merge(match_data, info[, .(match_no,date,team1,team2)], all.x = T, by = "match_no")
    all_matches <- match_data[, unique(match_no)]
    for(i in all_matches){
        fwrite(match_data[match_no == i, .SD, .SDcols = -"match_no"], file = paste0("./match_data_scraped/",info[match_no == i, strftime(date, "%Y-%m-%d")],"-",
                                                                       info[match_no == i,team1],"-",info[match_no == i,team2],".csv"))
    }
}