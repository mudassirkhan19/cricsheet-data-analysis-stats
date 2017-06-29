loadData <- function(){
    source('./Functions/read_cricsheet_csv.R')
    # Read IPL
    ipl_list <- readCricsheetCsv("./match_data/IPL")
    match_data_ipl <- ipl_list[[1]]
    match_data_ipl[, comp := "IPL"]
    info_ipl <- ipl_list[[2]]
    source('./Functions/matchInfoReshape.R')
    info_ipl <- matchInfoReshape(info_ipl, add_miss = TRUE)
    info_ipl[, comp := "IPL"]
    
    super_over_data_ipl <- match_data_ipl[grep("super over", L3, ignore.case = T)]
    match_data_ipl <- match_data_ipl[L3 == "1st innings" | L3 == "2nd innings"]
    temp <- match_data_ipl[, max(match_no)]
    
    # Read T20
    T20_list <- readCricsheetCsv("./match_data/T20", start_no = temp)
    match_data_T20 <- T20_list[[1]]
    match_data_T20[, comp := "T20"]
    info_T20 <- T20_list[[2]]
    info_T20 <- matchInfoReshape(info_T20)
    info_T20[, comp := "T20"]
    
    super_over_data_T20 <- match_data_T20[grep("super over", L3, ignore.case = T)]
    match_data_T20 <- match_data_T20[L3 == "1st innings" | L3 == "2nd innings"]
    temp <- match_data_T20[, max(match_no)]
    
    # Read BBL
    bbl_list <- readCricsheetCsv("./match_data/BBL", start_no = temp)
    match_data_bbl <- bbl_list[[1]]
    match_data_bbl[, comp := "BBL"]
    info_bbl <- bbl_list[[2]]
    info_bbl <- matchInfoReshape(info_bbl)
    info_bbl[, comp := "BBL"]
    
    super_over_data_bbl <- match_data_bbl[grep("super over", L3, ignore.case = T)]
    match_data_bbl <- match_data_bbl[L3 == "1st innings" | L3 == "2nd innings"]
    
    # Merge all lists
    match_data <- rbindlist(list(match_data_ipl, match_data_bbl, match_data_T20), fill = T)
    info <- rbindlist(list(info_ipl, info_bbl, info_T20), fill = T)
    super_over_data <- rbindlist(list(super_over_data_ipl, super_over_data_bbl, super_over_data_T20), fill = T)
    
    changecols <- setdiff(names(match_data),c("L3","batsman_NA","bowler_NA","non_striker_NA","wicket_fielders","wicket_kind","wicket_player_out"))
    match_data[, (changecols) := lapply(.SD, as.integer), .SDcols = changecols]
    super_over_data[, (changecols) := lapply(.SD, as.integer), .SDcols = changecols]
    match_data <- match_data[order(match_no, L2, over, ball)]
    
    bowler_dismissal <- setdiff(unique(match_data$wicket_kind), c("obstructing the field","run out","retired hurt", "hit the ball twice", "handled the ball", "timed out", NA))
    
    team_innings1 <- info[!is.na(toss_decision), .(team = toss, innings = ifelse(toss_decision == "field", 2, 1)), by = match_no]
    team_innings2 <- info[!is.na(toss_decision), .(team = setdiff(c(team1,team2),toss)), by = match_no]
    team_innings2$innings <- ifelse(team_innings1$innings == 1, 2, 1)
    team_innings <- rbind(team_innings1, team_innings2)[order(match_no, innings)]
    
    opposing_team <- copy(team_innings)
    opposing_team[, innings := ifelse(innings == 1, 2, 1)]
    
    return(list(match_data, info, team_innings, opposing_team, bowler_dismissal))
}