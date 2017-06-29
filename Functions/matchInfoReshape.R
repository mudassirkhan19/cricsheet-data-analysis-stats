## Cricsheet match Info Reshaper
## Make sure to include missing match info in "./csv_files/"

matchInfoReshape <- function(infoFile, add_miss = FALSE){
    require(data.table)
    require(reshape2)
    new_info <- infoFile[,.(match_no = unique(match_no))]
    dates <- infoFile[L2 == "dates", .(date = as.POSIXct(value, format = "%Y-%m-%d")), by = match_no]
    # picking later date in a multi day match
    dates <- dates[!shift(match_no, type = "lead") == match_no | match_no == length(unique(match_no))]
    cities <- infoFile[L2 == "city", .(match_no, city = value)]
    venues <- infoFile[L2 == "venue", .(match_no, venue = value)]
    neutral_venues <- infoFile[L2 == "neutral_venue", .(match_no, neutral_venue = value)]
    all_teams <- infoFile[L2 == "teams",.(team1 = value[1], team2 = value[2]),by = match_no]
    win_team <- infoFile[L2 == "outcome" & L3 == "winner", .(match_no, won = value)]
    super_over_win <- infoFile[L3 == "eliminator" & L2 == "outcome", .(match_no, won = value)]
    win_team <- merge(win_team, super_over_win, all = TRUE)
    
    # Merging data
    
    new_info <- merge(new_info, dates, all = TRUE)
    new_info <- merge(new_info, cities, all = TRUE)
    new_info <- merge(new_info, venues,all = TRUE)
    new_info <- merge(new_info, neutral_venues,all = TRUE)
    new_info <- merge(new_info, all_teams,all = TRUE)
    new_info <- merge(new_info, win_team,all = TRUE)
    
    new_info[!is.na(won), lost := setdiff(c(team1,team2), won), by = match_no]
    
    outcome_type <- infoFile[L2 == "outcome" & L3 == "by", .(result_type = L4),by = match_no]
    super_over_result <- infoFile[L3 == "result" & L2 == "outcome", .(match_no, result_type = value)]
    no_result <- infoFile[value == "no result" & L2 == "outcome", .(match_no, result_type = value)]
    outcome_type <- merge(outcome_type, super_over_result, all = TRUE)
    outcome_type <- merge(outcome_type, no_result, all = TRUE)
    margin <- infoFile[L2 == "outcome" & !is.na(L4), .(match_no, margin = value)]
    
    MOM <- infoFile[L2 == "player_of_match", .(MOM = paste0(value, collapse="/")), by = match_no]
    toss_win <- infoFile[L2 == "toss" & L3 == "winner", .(toss = value), by = match_no]
    toss_decision <- infoFile[L2 == "toss" & L3 == "decision", .(toss_decision = value), by = match_no]
    umpires <- infoFile[L2 == "umpires",.(umpire1 = value[1], umpire2 = value[2]),by = match_no]
    
    new_info <- merge(new_info, outcome_type, all = TRUE)
    new_info <- merge(new_info, margin, all = TRUE)
    new_info <- merge(new_info, MOM, all = TRUE)
    new_info <- merge(new_info, toss_win, all = TRUE)
    new_info <- merge(new_info, toss_decision, all = TRUE)
    new_info <- merge(new_info, umpires, all = TRUE)
    
    if(add_miss){
        miss_match <- fread("./csv_files/missing_match_info.csv", na.strings = c("NA",""))
        miss_match[, date := as.POSIXct(date, format = "%m/%d/%Y")]
        new_info <- rbind(new_info, miss_match)
    }
    new_info
}

# temp <- data.table(umpires = union(info[,unique(umpire1)],info[,unique(umpire2)]))
# max.len = max(dim(temp)[1], length(info[, unique(venue)]))
# temp$venue <- c(info[, unique(venue)], rep(NA,max.len - length(info[, unique(venue)])))
# temp$team <- c(info[, union(unique(team1), unique(team2))], rep(NA,max.len - length(info[,union(unique(team1), unique(team2))])))
# temp$city <- c(info[, unique(city)], rep(NA,max.len - length(info[, unique(city)])))