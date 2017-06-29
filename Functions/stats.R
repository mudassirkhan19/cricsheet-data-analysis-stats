library(reshape2)
library(data.table)
library(lubridate)
library(openxlsx)
library(plyr)


### Batting Stats

#### Batting Stats overall

# No of innings batted
innings <- match_data[, .N,by = .(match_no,batsman_NA ,wicket_player_out,non_striker_NA)]
bat_innings <- innings[,.(player = union(union(unique(batsman_NA),unique(wicket_player_out)),unique(non_striker_NA))),by = match_no][!is.na(player),.(innings = .N),by=player]

# No of times dismissed
total_dismissed <- match_data[wicket_kind != "retired hurt", .(dismissals = .N),by = .(player = wicket_player_out)]

# Not outs
not_out <- merge(bat_innings,total_dismissed,all = TRUE)
not_out[is.na(dismissals), dismissals := 0]
not_out[, NO := innings - dismissals]

# Total runs
total_runs <- match_data[,.(runs = sum(runs_batsman)),by = .(player = batsman_NA)]

# Highest Score
all_score <- match_data[,.(score = sum(runs_batsman)), by = .(match_no,player = batsman_NA)]
zero_ball_score <- match_data[batsman_NA != wicket_player_out, .(score = 0), by = .(match_no,player = wicket_player_out)]
all_score <- merge(all_score,zero_ball_score,by = c("match_no","player"),all = TRUE)
all_score[, ':=' (score = ifelse(is.na(score.x),0,score.x), score.x = NULL, score.y = NULL)]
high_score <- all_score[, .(HS = max(score)), by = player]

# Average
average <- merge(total_runs,total_dismissed,all = TRUE)
average[,average := ifelse(is.na(runs/dismissals), NA, format(round(runs/dismissals,2),nsmall = 2))]

# Balls Faced
balls_faced <- match_data[is.na(extras_wides) , .(BF = .N),by = .(player = batsman_NA)]

# Strike Rate
strike_rate <- merge(total_runs,balls_faced,all = TRUE)
strike_rate[, SR := ifelse(is.na(runs/BF), NA, format(round(runs/BF * 100,2),nsmall = 2)), by = player]

# Dots
dots <- match_data[is.na(extras_wides) & runs_batsman == 0, .(Dots = .N), by = .(player = batsman_NA)]

# No 0f 100's,50's,30's & 0's
all_score[,':='("100" = score>99, "50" = score>49, "30" = score>29, "0" = score==0)]
scores <- all_score[,.("100" = sum(`100`), "50" = sum(`50`), "30" = sum(`30`), "0" = sum(`0`)), by = player]

# boundaries
fours <- match_data[runs_batsman == 4 & is.na(runs_non_boundary), .('4s' = .N), by = .(player = batsman_NA)]
sixes <- match_data[runs_batsman == 6 & is.na(runs_non_boundary), .('6s' = .N), by = .(player = batsman_NA)]


# Combining above stats
batting_stats <- not_out[, .SD, .SDcols = -"dismissals"]
batting_stats <- merge(batting_stats, total_runs, all = TRUE)
batting_stats <- merge(batting_stats, high_score, all = TRUE)
batting_stats <- merge(batting_stats, average[,.(player,average)],all = TRUE)
batting_stats <- merge(batting_stats, balls_faced ,all = TRUE)
batting_stats <- merge(batting_stats, strike_rate[,.(player,SR)],all = TRUE)
batting_stats <- merge(batting_stats, scores, all = TRUE)
batting_stats <- merge(batting_stats, fours, all = TRUE)
batting_stats <- merge(batting_stats, sixes, all = TRUE)
batting_stats <- merge(batting_stats, dots, all = TRUE)
batting_stats[is.na(`4s`), `4s` := 0]
batting_stats[is.na(`6s`), `6s` := 0]

# Balls per boundary
batting_stats[, BpB := BF/(`4s`+`6s`)]
batting_stats[, Bp4 := BF/`4s`]
batting_stats[, Bp6 := BF/`6s`]

# Percentage Boundary Scores
batting_stats[, PrB := ((4 * `4s` + 6 * `6s`)/runs) * 100]
batting_stats[, Pr4 := ((4 * `4s`)/runs) * 100]
batting_stats[, Pr6 := ((6 * `6s`)/runs) * 100]

# Boundary Ball%
batting_stats[, BBP := (`4s`+`6s`)/BF * 100]

# Runs per boundary & non-boundary
batting_stats[, RpB := (4 * `4s` + 6 * `6s`)/(`4s`+`6s`)]
batting_stats[, RpNB := (runs - (4 * `4s` + 6 * `6s`))/(BF - (`4s`+`6s`))]

# Runs & Balls per innings
batting_stats[, RpI := runs/innings]
batting_stats[, BpI := BF/innings]

# Final Formatting
batting_stats[`100` == 0, `100` := NA]
batting_stats[`50` == 0, `50` := NA]
batting_stats[`30` == 0, `30` := NA]
batting_stats[`0` == 0, `0` := NA]
batting_stats[`4s` == 0, `4s` := NA]
batting_stats[`6s` == 0, `6s` := NA]
batting_stats[is.infinite(BpB), BpB := NA]
batting_stats[is.infinite(Bp4), Bp4 := NA]
batting_stats[is.infinite(Bp6), Bp6 := NA]
batting_stats[is.nan(PrB), PrB := NA]
batting_stats[is.nan(Pr4), Pr4 := NA]
batting_stats[is.nan(Pr6), Pr6 := NA]
batting_stats[is.nan(RpB), RpB := NA]

modify_excel_sheet(sheetName = "Overall_stats", wbPath = "./stats/Batsman_Stats.xlsx", data = batting_stats)



## Batsman stats by year


temp <- data.table()
temp1 <- info[, .(match_no, Season = year(date))]
match_played_data <- merge(match_data, temp1, by = "match_no")

innings <- match_played_data[, .(Innings = length(unique(match_no))), by = .(Batsman = batsman_NA, Season)]


wickets <- match_played_data[wicket_kind != "retired hurt", .(Wkts = .N), by = .(Batsman = wicket_player_out, Season)]

runs <- match_played_data[, .(Runs = sum(runs_batsman)), by = .(Batsman = batsman_NA, Season)]

balls <- match_played_data[is.na(extras_wides), .(BF = .N), by = .(Batsman = batsman_NA, Season)]

no_balls <- match_played_data[!is.na(extras_noballs), .(NoBl = .N), by = .(Batsman = batsman_NA, Season)]

all_score <- match_played_data[,.(score = sum(runs_batsman)), by = .(match_no,player = batsman_NA, Season)]
zero_ball_score <- match_played_data[batsman_NA != wicket_player_out, .(score = 0), by = .(match_no,player = wicket_player_out, Season)]
all_score <- merge(all_score,zero_ball_score,by = c("match_no","player","Season"),all = TRUE)
all_score[, ':=' (score = ifelse(is.na(score.x),0,score.x), score.x = NULL, score.y = NULL)]
high_score <- all_score[, .(HS = max(score)), by = .(Batsman = player, Season)]

outs <- match_played_data[wicket_kind != "retired hurt", .(out = TRUE), by = .(match_no,player=wicket_player_out,Season)]
all_score <- merge(outs, all_score, all = TRUE, by = c("match_no","player","Season"))
all_score[,':='("100" = score>99, "50" = score>49, "30" = score>29, "duck" = score == 0 & !is.na(out))]
scores <- all_score[,.("100" = sum(`100`), "50" = sum(`50`), "30" = sum(`30`), "duck" = sum(duck)), by = .(Batsman = player,Season)]

temp <- merge(innings, wickets, all = TRUE)
temp <- merge(temp, runs, all = TRUE)
temp <- merge(temp, balls, all = TRUE)
temp <- merge(temp, no_balls, all = TRUE)
temp <- merge(temp, high_score, all = TRUE)
temp <- merge(temp, scores, all = TRUE)

temp[, Avg := ifelse(is.na(Runs/Wkts), NA, format(round(Runs/Wkts,2),nsmall = 2))]
temp[, SR := ifelse(is.na(Runs/BF), NA, format(round(Runs/BF * 100,2),nsmall = 2))]

fours <- match_played_data[runs_batsman == 4 & is.na(runs_non_boundary), .('4s' = .N), by = .(Batsman = batsman_NA, Season)]

sixes <- match_played_data[runs_batsman == 6 & is.na(runs_non_boundary), .('6s' = .N), by = .(Batsman = batsman_NA, Season)]

temp <- merge(temp, fours, all = TRUE)
temp <- merge(temp, sixes, all = TRUE)

dots <- match_played_data[is.na(extras_wides) & runs_batsman == 0, .(Dots = .N), by = .(Batsman = batsman_NA, Season)]
temp <- merge(temp, dots, all = TRUE)

temp[, NoBl := NULL]

rm(list = c("innings","wickets","runs","balls","no_balls","fours","sixes","dots","match_played_data","zero_ball_score","all_score","high_score","scores","outs"))

modify_excel_sheet(sheetName = "Yearwise_stats", wbPath = "./stats/Batsman_Stats.xlsx", data = temp)


## Batsman Vs Teams

temp <- data.table()

match_played_data <- merge(match_data, opposing_team[, .(match_no, innings, Team = team)], by.x = c("match_no", "L2"), by.y = c("match_no", "innings"))

innings <- match_played_data[, .(Innings = length(unique(match_no))), by = .(Batsman = batsman_NA, Team)]

wickets <- match_played_data[wicket_kind != "retired hurt", .(Wkts = .N),by = .(Batsman = wicket_player_out, Team)]

runs <- match_played_data[, .(Runs = sum(runs_batsman)), by = .(Batsman = batsman_NA, Team)]

balls <- match_played_data[is.na(extras_wides), .(BF = .N), by = .(Batsman = batsman_NA, Team)]

no_balls <- match_played_data[!is.na(extras_noballs), .(NoBl = .N), by = .(Batsman = batsman_NA, Team)]

all_score <- match_played_data[,.(score = sum(runs_batsman)), by = .(match_no,player = batsman_NA, Team)]
zero_ball_score <- match_played_data[batsman_NA != wicket_player_out, .(score = 0), by = .(match_no,player = wicket_player_out, Team)]
all_score <- merge(all_score,zero_ball_score,by = c("match_no","player","Team"),all = TRUE)
all_score[, ':=' (score = ifelse(is.na(score.x),0,score.x), score.x = NULL, score.y = NULL)]
high_score <- all_score[, .(HS = max(score)), by = .(Batsman = player, Team)]

outs <- match_played_data[wicket_kind != "retired hurt", .(out = TRUE), by = .(match_no,player=wicket_player_out,Team)]
all_score <- merge(outs, all_score, all = TRUE, by = c("match_no","player","Team"))
all_score[,':='("100" = score>99, "50" = score>49, "30" = score>29, "duck" = score == 0 & !is.na(out))]
scores <- all_score[,.("100" = sum(`100`), "50" = sum(`50`), "30" = sum(`30`), "duck" = sum(duck)), by = .(Batsman = player,Team)]

temp <- merge(innings, wickets, all = TRUE)
temp <- merge(temp, runs, all = TRUE)
temp <- merge(temp, balls, all = TRUE)
temp <- merge(temp, no_balls, all = TRUE)
temp <- merge(temp, high_score, all = TRUE)
temp <- merge(temp, scores, all = TRUE)

temp[, Avg := ifelse(is.na(Runs/Wkts), NA, format(round(Runs/Wkts,2),nsmall = 2))]
temp[, SR := ifelse(is.na(Runs/BF), NA, format(round(Runs/BF * 100,2),nsmall = 2))]

fours <- match_played_data[runs_batsman == 4 & is.na(runs_non_boundary), .('4s' = .N), by = .(Batsman = batsman_NA, Team)]

sixes <- match_played_data[runs_batsman == 6 & is.na(runs_non_boundary), .('6s' = .N), by = .(Batsman = batsman_NA, Team)]

temp <- merge(temp, fours, all = TRUE)
temp <- merge(temp, sixes, all = TRUE)

dots <- match_played_data[is.na(extras_wides) & runs_batsman == 0, .(Dots = .N), by = .(Batsman = batsman_NA, Team)]
temp <- merge(temp, dots, all = TRUE)

temp[, Bowl_SR := ifelse(is.na(rowSums(.SD, na.rm = T)/Wkts), NA, format(round(rowSums(.SD, , na.rm = T)/Wkts,2),nsmall = 2)), .SDcols = c("BF","NoBl")]

temp[, NoBl := NULL]

rm(list = c("innings","wickets","runs","balls","no_balls","fours","sixes","dots","match_played_data","zero_ball_score","all_score","high_score","scores","outs"))

modify_excel_sheet(sheetName = "BatsmanVsTeam", wbPath = "./stats/Batsman_Stats.xlsx", data = temp)

## Batsman at Stadium Vs Teams

temp <- data.table()

opposing_team_venue <- merge(opposing_team[, .(match_no, innings, Team = team)], info[,.(match_no, Venue = venue)], all.x = TRUE, by = "match_no")

match_played_data <- merge(match_data, opposing_team_venue, by.x = c("match_no", "L2"), by.y = c("match_no", "innings"))

innings <- match_played_data[, .(Innings = length(unique(match_no))), by = .(Batsman = batsman_NA, Team, Venue)]


wickets <- match_played_data[wicket_kind != "retired hurt", .(Wkts = .N), by = .(Batsman = wicket_player_out, Team, Venue)]

runs <- match_played_data[, .(Runs = sum(runs_batsman)), by = .(Batsman = batsman_NA, Team, Venue)]

balls <- match_played_data[is.na(extras_wides), .(BF = .N), by = .(Batsman = batsman_NA, Team, Venue)]

no_balls <- match_played_data[!is.na(extras_noballs), .(NoBl = .N), by = .(Batsman = batsman_NA, Team, Venue)]

all_score <- match_played_data[,.(score = sum(runs_batsman)), by = .(match_no,player = batsman_NA, Team, Venue)]
zero_ball_score <- match_played_data[batsman_NA != wicket_player_out, .(score = 0), by = .(match_no,player = wicket_player_out, Team, Venue)]
all_score <- merge(all_score,zero_ball_score,by = c("match_no","player","Team","Venue"),all = TRUE)
all_score[, ':=' (score = ifelse(is.na(score.x),0,score.x), score.x = NULL, score.y = NULL)]
high_score <- all_score[, .(HS = max(score)), by = .(Batsman = player, Team, Venue)]

outs <- match_played_data[wicket_kind != "retired hurt", .(out = TRUE), by = .(match_no,player=wicket_player_out, Team, Venue)]
all_score <- merge(outs, all_score, all = TRUE, by = c("match_no","player","Team", "Venue"))
all_score[,':='("100" = score>99, "50" = score>49, "30" = score>29, "duck" = score == 0 & !is.na(out))]
scores <- all_score[, .("100" = sum(`100`), "50" = sum(`50`), "30" = sum(`30`), "duck" = sum(duck)), by = .(Batsman = player, Team, Venue)]

temp <- merge(innings, wickets, all = TRUE)
temp <- merge(temp, runs, all = TRUE)
temp <- merge(temp, balls, all = TRUE)
temp <- merge(temp, no_balls, all = TRUE)
temp <- merge(temp, high_score, all = TRUE)
temp <- merge(temp, scores, all = TRUE)

temp[, Avg := ifelse(is.na(Runs/Wkts), NA, format(round(Runs/Wkts,2),nsmall = 2))]
temp[, SR := ifelse(is.na(Runs/BF), NA, format(round(Runs/BF * 100,2),nsmall = 2))]

fours <- match_played_data[runs_batsman == 4 & is.na(runs_non_boundary), .('4s' = .N), by = .(Batsman = batsman_NA, Team, Venue)]

sixes <- match_played_data[runs_batsman == 6 & is.na(runs_non_boundary), .('6s' = .N), by = .(Batsman = batsman_NA, Team, Venue)]

temp <- merge(temp, fours, all = TRUE)
temp <- merge(temp, sixes, all = TRUE)

dots <- match_played_data[is.na(extras_wides) & runs_batsman == 0, .(Dots = .N), by = .(Batsman = batsman_NA, Team ,Venue)]
temp <- merge(temp, dots, all = TRUE)

temp[, Bowl_SR := ifelse(is.na(rowSums(.SD, na.rm = T)/Wkts), NA, format(round(rowSums(.SD, , na.rm = T)/Wkts,2),nsmall = 2)), .SDcols = c("BF","NoBl")]

temp[, NoBl := NULL]

rm(list = c("innings","wickets","runs","balls","no_balls","fours","sixes","dots","match_played_data","zero_ball_score","all_score","high_score","scores","outs"))

modify_excel_sheet(sheetName = "BatsmanVsTeam_stadium", wbPath = "./stats/Batsman_Stats.xlsx", data = temp)

## Batting Position

match_played_data <- merge(match_data, info[, .(match_no, Season = year(date))], by = "match_no")
temp1 <- c(1L,2L,3L,4L,5L,6L,7L,8L,9L,10L,11L)
match_played_data[, bat_pos := mapvalues(batsman_NA, unique(c(batsman_NA, non_striker_NA)), temp1[1:length(unique(c(batsman_NA, non_striker_NA)))]), by = .(match_no, L2)]
match_played_data[, NS_pos := mapvalues(non_striker_NA, unique(c(batsman_NA, non_striker_NA)), temp1[1:length(unique(c(batsman_NA, non_striker_NA)))]), by = .(match_no, L2)]

temp <- data.table()

innings <- match_played_data[, .(Str = .N),by = .(match_no,Season,bat_pos, Batsman = batsman_NA)]
temp3 <- match_played_data[, .(NS = .N),by = .(match_no,Season,bat_pos = NS_pos, Batsman = non_striker_NA)]
innings <- merge(innings, temp3, all = T)
innings <- innings[, .(innings = .N),by = .(Batsman,Season,bat_pos)]

temp2 <- merge(match_data, info[, .(match_no, Season = year(date))], by = "match_no")
temp1 <- c(1L,2L,3L,4L,5L,6L,7L,8L,9L,10L,11L)
temp2[, out_pos := mapvalues(wicket_player_out, unique(c(batsman_NA, non_striker_NA)), temp1[1:length(unique(c(batsman_NA, non_striker_NA)))]), by = .(match_no, L2)]


wickets <- temp2[wicket_kind != "retired hurt", .(Wkts = .N), by = .(Batsman = wicket_player_out, bat_pos = out_pos, Season)]

runs <- match_played_data[, .(Runs = sum(runs_batsman)), by = .(Batsman = batsman_NA, bat_pos, Season)]

balls <- match_played_data[is.na(extras_wides), .(BF = .N), by = .(Batsman = batsman_NA, bat_pos, Season)]

no_balls <- match_played_data[!is.na(extras_noballs), .(NoBl = .N), by = .(Batsman = batsman_NA, bat_pos, Season)]

all_score <- match_played_data[,.(score = sum(runs_batsman)), by = .(match_no,player = batsman_NA, Season, bat_pos)]
zero_ball_score <- match_played_data[, .(score = 0), by = .(match_no,player = non_striker_NA, Season, bat_pos = NS_pos)]
all_score <- merge(all_score,zero_ball_score,by = c("match_no","player","Season","bat_pos"),all = TRUE)
all_score[, ':=' (score = ifelse(is.na(score.x),0,score.x), score.x = NULL, score.y = NULL)]
high_score <- all_score[, .(HS = max(score)), by = .(Batsman = player, Season, bat_pos)]

outs <- temp2[wicket_kind != "retired hurt", .(out = TRUE), by = .(match_no,player=wicket_player_out, Season, bat_pos = out_pos)]
all_score <- merge(outs, all_score, all = TRUE, by = c("match_no","player","Season", "bat_pos"))
all_score[,':='("100" = score>99, "50" = score>49, "30" = score>29, "duck" = score == 0 & !is.na(out))]
scores <- all_score[, .("100" = sum(`100`), "50" = sum(`50`), "30" = sum(`30`), "duck" = sum(duck)), by = .(Batsman = player, Season, bat_pos)]

temp <- merge(innings, wickets, all = TRUE)
temp <- merge(temp, runs, all = TRUE)
temp <- merge(temp, balls, all = TRUE)
temp <- merge(temp, no_balls, all = TRUE)
temp <- merge(temp, high_score, all = TRUE)
temp <- merge(temp, scores, all = TRUE)

temp[, Avg := ifelse(is.na(Runs/Wkts), NA, format(round(Runs/Wkts,2),nsmall = 2))]
temp[, SR := ifelse(is.na(Runs/BF), NA, format(round(Runs/BF * 100,2),nsmall = 2))]

fours <- match_played_data[runs_batsman == 4 & is.na(runs_non_boundary), .('4s' = .N), by = .(Batsman = batsman_NA, bat_pos, Season)]

sixes <- match_played_data[runs_batsman == 6 & is.na(runs_non_boundary), .('6s' = .N), by = .(Batsman = batsman_NA, bat_pos, Season)]

temp <- merge(temp, fours, all = TRUE)
temp <- merge(temp, sixes, all = TRUE)

dots <- match_played_data[is.na(extras_wides) & runs_batsman == 0, .(Dots = .N), by = .(Batsman = batsman_NA, bat_pos, Season)]
temp <- merge(temp, dots, all = TRUE)

temp[, NoBl := NULL]

rm(list = c("innings","wickets","runs","balls","no_balls","fours","sixes","dots","match_played_data","zero_ball_score","all_score","high_score","outs","scores"))

modify_excel_sheet(sheetName = "bat_pos_stats", wbPath = "./stats/Batsman_Stats.xlsx", data = temp)


## Batsman matchwise score

match_played_data <- merge(match_data, team_innings, by.x = c("match_no","L2"), by.y = c("match_no","innings"),all = TRUE)
match_played_data <- merge(match_played_data, opposing_team[, .(match_no, opp_team = team, innings)], by.x = c("match_no","L2"), by.y = c("match_no","innings"),all = TRUE)
match_played_data <- merge(match_played_data, info[,.(match_no, won, Season = year(date))], by = "match_no",all = TRUE)
match_played_data[, won := ifelse(won == team, 1, 0)]

temp <- match_played_data[, .(Runs = sum(runs_batsman)), by = .(match_no, Inn = L2, Season, team, opp_team, batsman = batsman_NA)]
temp1 <- match_played_data[is.na(extras_wides), .(Balls = .N), by = .(match_no, Inn = L2, Season, team, opp_team, batsman = batsman_NA)]
temp2 <- match_played_data[is.na(runs_non_boundary) & runs_batsman == 4, .('4s' = .N), by = .(match_no, Inn = L2, Season, team, opp_team, batsman = batsman_NA)]
temp3 <- match_played_data[runs_batsman == 6 & is.na(runs_non_boundary), .('6s' = .N), by = .(match_no, Inn = L2, Season, team, opp_team, batsman = batsman_NA)]
temp4 <- match_played_data[is.na(extras_wides) & runs_batsman == 0, .(Dots = .N), by = .(match_no, Inn = L2, Season, team, opp_team, batsman = batsman_NA)]
temp5 <- match_played_data[, .(teamTotal = sum(runs_total)), by = .(match_no, team)]

temp <- merge(temp, temp1, all = TRUE)
temp <- merge(temp, temp2, all = TRUE)
temp <- merge(temp, temp3, all = TRUE)
temp <- merge(temp, temp4, all = TRUE)
temp[, SR := format(round((Runs/Balls * 100), 2),nsmall = 2)]
temp <- merge(temp, temp5, all = TRUE, by = c("match_no","team"))
temp[, `Score%` := format(round((Runs/teamTotal * 100), 2),nsmall = 2)]

modify_excel_sheet(sheetName = "matchwise_score", wbPath = "./stats/Batsman_Stats.xlsx", data = temp)

## Ballwise Stats Batsman

match_played_data <- merge(match_data, info[, .(match_no, Season = year(date))], by = "match_no")
setnames(opposing_team, "team","ag_team")
match_played_data <- merge(match_played_data, opposing_team, by.x = c("match_no","L2"), by.y = c("match_no","innings"))
setnames(opposing_team, "ag_team","team")

match_played_data[, bats_ball := cumsum(is.na(extras_wides)), by = .(match_no, batsman_NA)]
breaks <- c(0,5,10,15,20,25,30,35,40,45,50,55,60,65,70,75)
temp <- match_played_data[is.na(extras_wides), .(SR = mean(runs_batsman) * 100, Runs = sum(runs_batsman), Bdry_run = sum(runs_batsman[runs_batsman > 4 & is.na(runs_non_boundary)])), by = .(match_no, Season, batsman_NA, ag_team, five_b = cut(bats_ball, breaks))]
temp <- temp[order(match_no, ag_team, batsman_NA, five_b)]

modify_excel_sheet(sheetName = "matchwise_5balls_bat", wbPath = "./stats/Batsman_Stats.xlsx", data = temp)


## Batsman Dismissal Overwise

temp <- match_data[!is.na(wicket_player_out), .(dismissal = .N), by = .(batsman = wicket_player_out , over)]
temp[, Total := sum(dismissal), by = batsman]
temp <- dcast(temp, batsman + Total ~ over, value.var = "dismissal")

modify_excel_sheet(sheetName = "out_eachover", wbPath = "./stats/Batsman_Stats.xlsx", data = temp)

## Batsman Runs overwise

temp <- match_data[, .(runs = sum(runs_batsman)), by = .(batsman = batsman_NA, over)]
temp[, Total := sum(runs), by = batsman]
temp <- dcast(temp, batsman + Total ~ over, value.var = "runs")

modify_excel_sheet(sheetName = "runs_eachover", wbPath = "./stats/Batsman_Stats.xlsx", data = temp)

## Batsman Runs ballwise

temp <- match_data[, .(runs = sum(runs_batsman)), by = .(batsman = batsman_NA, ball)]
temp[, Total := sum(runs), by = batsman]
temp <- dcast(temp, batsman + Total ~ ball, value.var = "runs")

modify_excel_sheet(sheetName = "runs_eachball", wbPath = "./stats/Batsman_Stats.xlsx", data = temp)


### Bowling Stats

#### Bowling Stats overall

# Innings
bowl_innings <- match_data[, .N ,by = .(match_no,bowler_NA)][,.(innings = .N), by = .(player = bowler_NA)]

# No of Balls
balls_bowled <- match_data[is.na(extras_wides) & is.na(extras_noballs), .(Balls = .N), by = .(player = bowler_NA)]

# Over Data & maidens
rpo <- match_data[, .(over_runs = sum(runs_total)), by =.(match_no, bowler_NA, over)]
maidens <- rpo[over_runs == 0, .(maidens = .N), by = .(player = bowler_NA)]
rm(rpo)

# Runs Conceded
runs_conceded <- match_data[, .(Runs = sum(runs_total)-sum(extras_byes,na.rm = T)-sum(extras_legbyes,na.rm = T)), by = .(player = bowler_NA)]

# Wickets

total_wickets <- match_data[wicket_kind %in% bowler_dismissal, .(Wkts = .N), by = .(player = bowler_NA)]

# Combining stats
bowl_stats <- merge(bowl_innings, balls_bowled, all = TRUE)
bowl_stats <- merge(bowl_stats, maidens, all = TRUE)
bowl_stats <- merge(bowl_stats, runs_conceded, all = TRUE)
bowl_stats <- merge(bowl_stats, total_wickets, all = TRUE)

# Average
bowl_stats[, Ave := ifelse(is.na(Runs/Wkts), NA, format(round(Runs/Wkts,2),nsmall = 2))]

# Economy
bowl_stats[, Econ := ifelse(is.na(Runs/Balls), NA, format(round(Runs/(Balls/6),2),nsmall = 2))]

# Strike Rate
bowl_stats[, SR := ifelse(is.na(Balls/Wkts), NA, format(round(Balls/Wkts,2),nsmall = 2))]

# Wicket Per Innings
bowl_stats[, WPI := ifelse(is.na(Wkts/innings), NA, format(round(Wkts/innings,2),nsmall = 2))]

# Wicket Per 4overs
bowl_stats[, Wp4 := ifelse(is.na(Wkts/Balls), NA, format(round(Wkts/(Balls/24),2),nsmall = 2))]

# 0, 3w+, 4 & 5 Wickets
wickets_per_match <- match_data[wicket_kind %in% bowler_dismissal, .(Wkts = .N), by = .(match_no, player = bowler_NA)]
runs_per_match <- match_data[, .(runs_match = sum(runs_total)), by = .(match_no, player = bowler_NA)]
bowl_figures <- merge(wickets_per_match, runs_per_match, all = TRUE)
zero_wickets <- bowl_figures[is.na(Wkts), .("0w" = .N), by = player]
three_plus_wickets <- bowl_figures[Wkts == 3, .("3w+" = .N), by = player]
four_wickets <- bowl_figures[Wkts == 4, .("4w" = .N), by = player]
five_wickets <- bowl_figures[Wkts == 5, .("5w" = .N), by = player]

# Extras
wides <- match_data[, .(Wd = sum(!is.na(extras_wides))), by = .(player = bowler_NA)]

no_balls <- match_data[, .(Nb = sum(!is.na(extras_noballs))), by = .(player = bowler_NA)]

# Combining above stats
bowl_stats <- merge(bowl_stats, zero_wickets, all = TRUE)
bowl_stats <- merge(bowl_stats, three_plus_wickets, all = TRUE)
bowl_stats <- merge(bowl_stats, four_wickets, all = TRUE)
bowl_stats <- merge(bowl_stats, five_wickets, all = TRUE)
bowl_stats <- merge(bowl_stats, wides, all = TRUE)
bowl_stats <- merge(bowl_stats, no_balls, all = TRUE)

# Accuracy
bowl_stats[, Acc := ((Balls - (Nb + Wd))/Balls) * 100]

# Final Formatting
bowl_stats[Runs == 0, Runs := NA]
bowl_stats[Wd == 0, Wd := NA]
bowl_stats[Nb == 0, Nb := NA]

modify_excel_sheet(sheetName = "Overall_stats", wbPath = "./stats/Bowler_Stats.xlsx", data = bowl_stats)


## Bowler Stats by Year

temp <- data.table()
temp1 <- info[, .(match_no, Season = year(date))]
match_played_data <- merge(match_data, temp1, by = "match_no")

innings <- match_played_data[, .(Innings = length(unique(match_no))), by = .(Bowler = bowler_NA, Season)]


runs <- match_played_data[, .(Runs = sum(runs_total)-sum(extras_byes,na.rm = T)-sum(extras_legbyes,na.rm = T)), by = .(Bowler = bowler_NA, Season)]

balls <- match_played_data[is.na(extras_wides) & is.na(extras_noballs), .(Balls = .N), by = .(Bowler = bowler_NA, Season)]

wickets <- match_played_data[wicket_kind %in% bowler_dismissal, .(Wkts = .N), by = .(Bowler = bowler_NA, Season)]

most_wickets <- match_played_data[, .(wickets = sum(!is.na(wicket_player_out)), Runs = sum(runs_total)-sum(extras_byes,na.rm = T)-sum(extras_legbyes,na.rm = T)), by = .(Bowler = bowler_NA,Season, match_no)][, .SD[wickets == max(wickets)], by = .(Bowler, Season)]
Best <- most_wickets[, .SD[which.min(Runs)], by = .(Bowler, Season)]
Best[, ':='(BF = paste(wickets, Runs,sep="/"), wickets=NULL, Runs = NULL, match_no = NULL)]

no_balls <- match_played_data[!is.na(extras_noballs), .(No = .N), by = .(Bowler = bowler_NA, Season)]

wides <- match_played_data[!is.na(extras_wides), .(Wd = .N), by = .(Bowler = bowler_NA, Season)]

wickets_per_match <- match_played_data[wicket_kind %in% bowler_dismissal, .(Wkts = .N), by = .(match_no, Season, Bowler = bowler_NA)]
runs_per_match <- match_played_data[, .(runs_match = sum(runs_total)), by = .(match_no, Season, Bowler = bowler_NA)]
bowl_figures <- merge(wickets_per_match, runs_per_match, all = TRUE)

three_plus_wickets <- bowl_figures[Wkts == 3, .("3w+" = .N), by = .(Bowler, Season)]
four_wickets <- bowl_figures[Wkts == 4, .("4w" = .N), by = .(Bowler, Season)]
five_wickets <- bowl_figures[Wkts == 5, .("5w" = .N), by = .(Bowler, Season)]

temp <- merge(innings, runs, all = TRUE)
temp <- merge(temp, balls, all = TRUE)
temp <- merge(temp, wickets, all = TRUE)
temp <- merge(temp, Best, all = TRUE)
temp <- merge(temp, no_balls, all = TRUE)
temp <- merge(temp, wides, all = TRUE)

temp <- merge(temp, three_plus_wickets, all = TRUE)
temp <- merge(temp, four_wickets, all = TRUE)
temp <- merge(temp, five_wickets, all = TRUE)

dots <- match_played_data[runs_batsman == 0, .(Dots = .N), by = .(Bowler = bowler_NA, Season)]
temp <- merge(temp, dots, all = TRUE)

temp[!is.na(Wkts), Avg := format(round(Runs/Wkts,2),nsmall = 2)]
temp[!is.na(Wkts), SR := format(round(Balls/Wkts,2),nsmall = 2)]

rm(list = c("innings","wickets","runs","balls","no_balls","wides","most_wickets","Best","four_wickets","five_wickets","three_plus_wickets","dots","match_played_data"))

modify_excel_sheet(sheetName = "YearWise_stats", wbPath = "./stats/Bowler_Stats.xlsx", data = temp)


# Bowler Vs Teams

temp <- data.table()

match_played_data <- merge(match_data, team_innings, by.x = c("match_no", "L2"), by.y = c("match_no", "innings"))
setnames(match_played_data, "team","Team")

innings <- match_played_data[, .(Innings = length(unique(match_no))), by = .(Bowler = bowler_NA, Team)]


runs <- match_played_data[, .(Runs = sum(runs_total)-sum(extras_byes,na.rm = T)-sum(extras_legbyes,na.rm = T)), by = .(Bowler = bowler_NA, Team)]

balls <- match_played_data[is.na(extras_wides) & is.na(extras_noballs), .(Balls = .N), by = .(Bowler = bowler_NA, Team)]

wickets <- match_played_data[wicket_kind %in% bowler_dismissal, .(Wkts = .N), by = .(Bowler = bowler_NA, Team)]

most_wickets <- match_played_data[, .(wickets = sum(!is.na(wicket_player_out)), Runs = sum(runs_total)-sum(extras_byes,na.rm = T)-sum(extras_legbyes,na.rm = T)), by = .(Bowler = bowler_NA,Team, match_no)][, .SD[wickets == max(wickets)], by = .(Bowler, Team)]
Best <- most_wickets[, .SD[which.min(Runs)], by = .(Bowler, Team)]
Best[, ':='(BF = paste(wickets, Runs,sep="/"), wickets=NULL, Runs = NULL, match_no = NULL)]

no_balls <- match_played_data[!is.na(extras_noballs), .(No = .N), by = .(Bowler = bowler_NA, Team)]

wides <- match_played_data[!is.na(extras_wides), .(Wd = .N), by = .(Bowler = bowler_NA, Team)]

wickets_per_match <- match_played_data[wicket_kind %in% bowler_dismissal, .(Wkts = .N), by = .(match_no, Team, Bowler = bowler_NA)]
runs_per_match <- match_played_data[, .(runs_match = sum(runs_total)), by = .(match_no, Team, Bowler = bowler_NA)]
bowl_figures <- merge(wickets_per_match, runs_per_match, all = TRUE)

three_plus_wickets <- bowl_figures[Wkts == 3, .("3w+" = .N), by = .(Bowler, Team)]
four_wickets <- bowl_figures[Wkts == 4, .("4w" = .N), by = .(Bowler, Team)]
five_wickets <- bowl_figures[Wkts == 5, .("5w" = .N), by = .(Bowler, Team)]

temp <- merge(innings, runs, all = TRUE)
temp <- merge(temp, balls, all = TRUE)
temp <- merge(temp, wickets, all = TRUE)
temp <- merge(temp, Best, all = TRUE)
temp <- merge(temp, no_balls, all = TRUE)
temp <- merge(temp, wides, all = TRUE)

temp <- merge(temp, three_plus_wickets, all = TRUE)
temp <- merge(temp, four_wickets, all = TRUE)
temp <- merge(temp, five_wickets, all = TRUE)

dots <- match_played_data[runs_batsman == 0, .(Dots = .N), by = .(Bowler = bowler_NA, Team)]
temp <- merge(temp, dots, all = TRUE)

temp[!is.na(Wkts), Avg := format(round(Runs/Wkts,2),nsmall = 2)]
temp[!is.na(Wkts), SR := format(round(Balls/Wkts,2),nsmall = 2)]

rm(list = c("innings","wickets","runs","balls","no_balls","wides","most_wickets","Best","four_wickets","five_wickets","three_plus_wickets","dots","match_played_data"))

modify_excel_sheet(sheetName = "BowlerVsTeam", wbPath = "./stats/Bowler_Stats.xlsx", data = temp)


# Bowler Vs Teams at stadiums

temp <- data.table()
temp1 <- merge(team_innings, info[,.(match_no, venue)], by = "match_no", all.x = TRUE)

match_played_data <- merge(match_data, temp1, by.x = c("match_no", "L2"), by.y = c("match_no", "innings"))
setnames(match_played_data, c("team","venue"),c("Team","Venue"))

innings <- match_played_data[, .(Innings = length(unique(match_no))), by = .(Bowler = bowler_NA, Team, Venue)]


runs <- match_played_data[, .(Runs = sum(runs_total)-sum(extras_byes,na.rm = T)-sum(extras_legbyes,na.rm = T)), by = .(Bowler = bowler_NA, Team, Venue)]

balls <- match_played_data[is.na(extras_wides) & is.na(extras_noballs), .(Balls = .N), by = .(Bowler = bowler_NA, Team, Venue)]

wickets <- match_played_data[wicket_kind %in% bowler_dismissal, .(Wkts = .N), by = .(Bowler = bowler_NA, Team, Venue)]

most_wickets <- match_played_data[, .(wickets = sum(!is.na(wicket_player_out)), Runs = sum(runs_total)-sum(extras_byes,na.rm = T)-sum(extras_legbyes,na.rm = T)), by = .(Bowler = bowler_NA,Team, Venue, match_no)][, .SD[wickets == max(wickets)], by = .(Bowler, Team, Venue)]
Best <- most_wickets[, .SD[which.min(Runs)], by = .(Bowler, Team, Venue)]
Best[, ':='(BF = paste(wickets, Runs,sep="/"), wickets=NULL, Runs = NULL, match_no = NULL)]

no_balls <- match_played_data[!is.na(extras_noballs), .(No = .N), by = .(Bowler = bowler_NA, Team, Venue)]

wides <- match_played_data[!is.na(extras_wides), .(Wd = .N), by = .(Bowler = bowler_NA, Team, Venue)]

wickets_per_match <- match_played_data[wicket_kind %in% bowler_dismissal, .(Wkts = .N), by = .(match_no, Team, Venue, Bowler = bowler_NA)]
runs_per_match <- match_played_data[, .(runs_match = sum(runs_total)), by = .(match_no, Team, Venue, Bowler = bowler_NA)]
bowl_figures <- merge(wickets_per_match, runs_per_match, all = TRUE)

three_plus_wickets <- bowl_figures[Wkts == 3, .("3w+" = .N), by = .(Bowler, Team, Venue)]
four_wickets <- bowl_figures[Wkts == 4, .("4w" = .N), by = .(Bowler, Team, Venue)]
five_wickets <- bowl_figures[Wkts == 5, .("5w" = .N), by = .(Bowler, Team, Venue)]

temp <- merge(innings, runs, all = TRUE)
temp <- merge(temp, balls, all = TRUE)
temp <- merge(temp, wickets, all = TRUE)
temp <- merge(temp, Best, all = TRUE)
temp <- merge(temp, no_balls, all = TRUE)
temp <- merge(temp, wides, all = TRUE)

temp <- merge(temp, three_plus_wickets, all = TRUE)
temp <- merge(temp, four_wickets, all = TRUE)
temp <- merge(temp, five_wickets, all = TRUE)

dots <- match_played_data[runs_batsman == 0, .(Dots = .N), by = .(Bowler = bowler_NA, Team, Venue)]
temp <- merge(temp, dots, all = TRUE)

temp[!is.na(Wkts), Avg := format(round(Runs/Wkts,2),nsmall = 2)]
temp[!is.na(Wkts), SR := format(round(Balls/Wkts,2),nsmall = 2)]

rm(list = c("innings","wickets","runs","balls","no_balls","wides","most_wickets","Best","four_wickets","five_wickets","three_plus_wickets","dots","match_played_data"))

modify_excel_sheet(sheetName = "BowlerVsTeam_stadium", wbPath = "./stats/Bowler_Stats.xlsx", data = temp)


## Bowler matchwise stats

temp <- match_data[is.na(extras_wides) & is.na(extras_noballs), .(Balls = .N), by = .(match_no, Player = bowler_NA)]

temp[, O := paste(Balls %/% 6, Balls %% 6, sep = ".")]


rpo <- match_data[, .(over_runs = sum(runs_total)), by =.(match_no, bowler_NA, over)]
temp1 <- rpo[over_runs == 0, .(M = .N), by = .(match_no, Player = bowler_NA)]
rm(rpo)

temp2 <- match_data[, .(R = sum(runs_total)-sum(extras_byes,na.rm = T)-sum(extras_legbyes,na.rm = T)), by = .(match_no, Player = bowler_NA)]

temp3 <- match_data[!is.na(wicket_player_out), .(W = .N), by = .(match_no, Player = bowler_NA)]

temp <- merge(temp, temp1, all = TRUE)
temp <- merge(temp, temp2, all = TRUE)
temp <- merge(temp, temp3, all = TRUE)
temp[, Econ := ifelse(is.na(R/Balls), NA, format(round(R/(Balls/6),2),nsmall = 2))]
temp[, Balls := NULL]

modify_excel_sheet(sheetName = "matchwise_figure", wbPath = "./stats/Bowler_Stats.xlsx", data = temp)


## Bowling Position

match_played_data <- merge(match_data, info[, .(match_no, Season = year(date))], by = "match_no")
temp <- match_played_data[, .(bowled = .N), by = .(player = bowler_NA, Season, match_no,over)][, .(bowled = .N), by = .(player, Season, over)]

temp1 <- match_played_data[, .N ,by = .(match_no, Season, bowler_NA)][,.(innings = .N), by = .(player = bowler_NA, Season)]
temp <- merge(temp, temp1, all = TRUE)

modify_excel_sheet(sheetName = "bowl_pos", wbPath = "./stats/Bowler_Stats.xlsx", data = temp)


## Bowler's Wickets batting positionwise

match_played_data <- merge(match_data, info[, .(match_no, Season = year(date))], by = "match_no")
temp1 <- c(1L,2L,3L,4L,5L,6L,7L,8L,9L,10L,11L)
match_played_data[, out_pos := mapvalues(wicket_player_out, unique(c(batsman_NA, non_striker_NA)), temp1[1:length(unique(c(batsman_NA, non_striker_NA)))]), by = .(match_no, L2)]

temp <- match_played_data[wicket_kind %in% bowler_dismissal, .(wickets = .N), by = .(bowler = bowler_NA , Season, out_pos)]
temp[, Total := sum(wickets), by = .(bowler, Season)]
temp <- dcast(temp, bowler + Season + Total ~ out_pos, value.var = "wickets")
setcolorder(temp, c("bowler","Season","1","2","3","4","5","6","7","8","9","10","11","Total"))

modify_excel_sheet(sheetName = "bat_pos", wbPath = "./stats/Bowler_Stats.xlsx", data = temp)


# Wickets Bowler's overwise

match_played_data <- merge(match_data, info[, .(match_no, Season = year(date))], by = "match_no")
temp1 <- c(1L,2L,3L,4L)
match_played_data[, bowl_over := mapvalues(over, unique(over)[order(unique(over))], temp1[1:length(unique(over))]), by = .(match_no, bowler_NA)]

temp <- match_played_data[wicket_kind %in% bowler_dismissal, .(wickets = .N), by = .(bowler = bowler_NA , Season, bowl_over)]
temp[, Total := sum(wickets), by = .(bowler, Season)]
temp <- dcast(temp, bowler + Season + Total ~ bowl_over, value.var = "wickets")

modify_excel_sheet(sheetName = "wkts_bowler_over", wbPath = "./stats/Bowler_Stats.xlsx", data = temp)


## Bowler's wickets overwise


temp <- match_data[wicket_kind %in% bowler_dismissal, .(wickets = .N), by = .(bowler = bowler_NA ,over)]
temp[, Total := sum(wickets), by = bowler]
temp <- dcast(temp, bowler + Total ~ over, value.var = "wickets")

modify_excel_sheet(sheetName = "wkts_over", wbPath = "./stats/Bowler_Stats.xlsx", data = temp)


## Runs Conceded overwise

temp <- match_data[, .(runs = sum(runs_total)-sum(extras_byes,na.rm = T)-sum(extras_legbyes,na.rm = T)), by = .(bowler = bowler_NA ,over)]
temp[, Total := sum(runs), by = bowler]
temp <- dcast(temp, bowler + Total ~ over, value.var = "runs")

modify_excel_sheet(sheetName = "runs_over", wbPath = "./stats/Bowler_Stats.xlsx", data = temp)



### Team Stats

## Team Wins and Losses

temp1 <- info[!is.na(won), .(Won = .N) , by = .(Team = won)]
temp2 <- info[!is.na(lost), .(Lost = .N) , by = .(Team = lost)]
temp3 <- data.table(info[, table(c(team1, team2))])
setnames(temp3, c("V1","N"), c("Team","Played"))
temp <- merge(temp3, temp1)
temp <- merge(temp,temp2)
temp[, NR := Played - (Won + Lost)]
temp <- temp[order(-Won)]

# Against runs and avgs

setnames(opposing_team, "team", "Team")

match_played_data <- merge(match_data, opposing_team, by.x = c("match_no","L2"), by.y = c("match_no","innings"))

NR_matches <- info[result_type == "no result", match_no]

temp1 <- match_played_data[!match_no %in% NR_matches, .(Runs = sum(runs_total)), by = .(Team, match_no, Innings = L2)]

match_played_data[,  balls := cumsum(is.na(extras_wides) & is.na(extras_noballs)), by = .(match_no, L2 ,over)]
match_played_data[!(is.na(extras_wides) & is.na(extras_noballs)), balls := balls + 1]

temp2 <- match_played_data[!match_no %in% NR_matches, .(Over = max(over), Ball = max(balls[over == max(over)])), by = .(Team, match_no, Innings = L2)]

temp1 <- merge(temp1, temp2)

temp1 <- temp1[, .(ag_RR = format(round(sum(Runs)/sum(Over-1, Ball/6), 2), nsmall = 2), ag_Max = max(Runs), ag_Min = min(Runs), ag_Avg = round(mean(Runs), 0)), by = Team]

temp <- merge(temp, temp1)

rm(match_played_data)
setnames(opposing_team, "Team", "team")

modify_excel_sheet(sheetName = "Win_loss", wbPath = "./stats/Team_Stats.xlsx", data = temp)

## Balls left in innings of completed matches

match_played_data <- merge(match_data, team_innings, by.x = c("match_no","L2"), by.y = c("match_no","innings"))
match_played_data <- merge(match_played_data, info[, .(match_no, won, result_type, Season = year(date))], by = "match_no")

match_played_data[,  ball := cumsum(is.na(extras_wides) & is.na(extras_noballs)), by = .(match_no, L2 ,over)]
match_played_data[!(is.na(extras_wides) & is.na(extras_noballs)), ball := ball + 1]

temp <- match_played_data[!is.na(won), .(Over = max(over), Ball = max(ball[over == max(over)])), by = .(match_no, Season, Inn = L2, Team = team, Won = (ifelse(team == won, 1, 0)), result_type)]
temp[, ':='(Balls_left = (6 * (20-Over) + (6-Ball)), Over = NULL, Ball = NULL)]

modify_excel_sheet(sheetName = "Balls_left_innings", wbPath = "./stats/Team_Stats.xlsx", data = temp)

## Team vs Team

temp <- info[, .(Played = .N), by = .(Team1 = pmin(team1, team2), Team2 = pmax(team1, team2))]

temp1 <- info[!is.na(won), .(Won1 = .N) , by = .(Team1 = pmin(team1, team2), Team2 = pmax(team1, team2), won)]
temp1[, Won2 := 0]
temp1[, ':='(Won1 = ifelse(Team1 == won, Won1, 0), Won2 = ifelse(Team2 == won, Won1, 0))]
temp1 <- temp1[, .(T1_wins = max(Won1), T2_wins = max(Won2)), by = .(Team1, Team2)]

temp <- merge(temp, temp1, all = TRUE)
temp[, ':='(Team = paste(Team1, Team2, sep = " vs "), NR = Played - rowSums(.SD, na.rm = T)), .SDcols = c("T1_wins","T2_wins")]
setcolorder(temp, c("Team","Team1","Team2","Played","T1_wins","T2_wins","NR"))

modify_excel_sheet(sheetName = "TeamVsTeam", wbPath = "./stats/Team_Stats.xlsx", data = temp)


## Team vs Team at Venue

temp <- info[, .(Played = .N), by = .(Team1 = pmin(team1, team2), Team2 = pmax(team1, team2), Venue = venue)]

temp1 <- info[!is.na(won), .(Won1 = .N) , by = .(Team1 = pmin(team1, team2), Team2 = pmax(team1, team2), venue, won)]
temp1[, Won2 := 0]
temp1[, ':='(Won1 = ifelse(Team1 == won, Won1, 0), Won2 = ifelse(Team2 == won, Won1, 0))]
temp1 <- temp1[, .(T1_wins = max(Won1), T2_wins = max(Won2)), by = .(Team1, Team2, Venue = venue)]

temp <- merge(temp, temp1, all = TRUE)
temp[, ':='(Team = paste(Team1, Team2, sep = " vs "), NR = Played - rowSums(.SD, na.rm = T)), .SDcols = c("T1_wins","T2_wins")]
setcolorder(temp, c("Team","Team1","Team2","Venue","Played","T1_wins","T2_wins","NR"))

modify_excel_sheet(sheetName = "TeamVsTeam_venue", wbPath = "./stats/Team_Stats.xlsx", data = temp)


## Teams Runs scored overwise

match_played_data <- merge(match_data, team_innings, by.x = c("match_no","L2"), by.y = c("match_no","innings"))

setnames(opposing_team, "team","field_team")
match_played_data <- merge(match_played_data, opposing_team, by.x = c("match_no","L2"), by.y = c("match_no","innings"))
setnames(opposing_team, "field_team","team")
match_played_data <- merge(match_played_data, info[, .(match_no, Season = year(date))], by = "match_no")

temp <- match_played_data[, .(Runs = sum(runs_total), Wkts = sum(!is.na(wicket_player_out))), by = .(match_no, Inn = L2, over, Season, bat_team = team, field_team)]
temp1 <- match_played_data[is.na(runs_non_boundary), .(`4s` = sum(runs_batsman == 4), `6s` = sum(runs_batsman == 6)), by = .(match_no, Inn = L2, over, Season, bat_team = team, field_team)]

temp <- merge(temp, temp1)
temp[, ':='(Ttl_runs = cumsum(Runs), Ttl_wkts = cumsum(Wkts), `Ttl_4s` = cumsum(`4s`), `Ttl_6s` = cumsum(`6s`)), by = .(match_no, Inn)]

modify_excel_sheet(sheetName = "overwise_score", wbPath = "./stats/Team_Stats.xlsx", data = temp)

## Toss Wins/Losses

temp1 <- info[!is.na(toss), .(TossWon = .N) , by = .(Team = toss)]
info[!is.na(toss), tossLost := ifelse(team1 == toss, team2, team1)]
temp2 <- info[!is.na(toss), .(TossLost = .N), by = .(Team = tossLost)]
info[, TWMW := ifelse(toss == won, 1, 0)]
info[, TLMW := ifelse(tossLost == won, 1, 0)]
temp3 <- info[, .(TWMW = sum(TWMW, na.rm = TRUE)), by = .(Team = toss)]
temp <- merge(temp1, temp2)
temp <- merge(temp, temp3)
temp[, TWML := (TossWon - TWMW)]
temp3 <- info[, .(TLMW = sum(TLMW, na.rm = TRUE)), by = .(Team = tossLost)]
temp <- merge(temp, temp3)
temp[, TLML := (TossLost - TLMW)]
temp[, totalToss := TossWon + TossLost]
setcolorder(temp, c("Team","totalToss","TossWon","TossLost","TWMW","TWML","TLMW","TLML"))

info[, c("tossLost","TWMW","TLMW") := NULL]

modify_excel_sheet(sheetName = "Toss_Win_Loss", wbPath = "./stats/Team_Stats.xlsx", data = temp)


# Toss Win/Loss by Venue

temp1 <- info[!is.na(toss), .(TossWon = .N) , by = .(Team = toss, Venue = venue)]
info[!is.na(toss), tossLost := ifelse(team1 == toss, team2, team1)]
temp2 <- info[!is.na(toss), .(TossLost = .N), by = .(Team = tossLost, Venue = venue)]
info[, TWMW := ifelse(toss == won, 1, 0)]
info[, TLMW := ifelse(tossLost == won, 1, 0)]
temp3 <- info[!is.na(toss), .(TWMW = sum(TWMW, na.rm = TRUE)), by = .(Team = toss, Venue = venue)]
temp <- merge(temp1, temp2, all = TRUE)
temp <- merge(temp, temp3, all = TRUE)
temp[, TWML := (TossWon - TWMW)]
temp3 <- info[!is.na(toss), .(TLMW = sum(TLMW, na.rm = TRUE)), by = .(Team = tossLost, Venue = venue)]
temp <- merge(temp, temp3, all = TRUE)
temp[, TLML := (TossLost - TLMW)]
info[, c("tossLost","TWMW","TLMW") := NULL]

modify_excel_sheet(sheetName = "Toss_Win_Loss_Venue", wbPath = "./stats/Team_Stats.xlsx", data = temp)


## Partnership Stats

## Overall Stats

temp <- match_data[, .(Inn = length(unique(match_no))), by = .(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA))]

temp1 <- match_data[, .(Runs1 = sum(runs_batsman)), by = .(batsman_NA, non_striker_NA)]
temp1[, Runs2 := 0]
temp1[, ':='(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), Runs1 = ifelse(batsman_NA == pmin(batsman_NA, non_striker_NA), Runs1,0),  Runs2 = ifelse(batsman_NA == pmax(batsman_NA, non_striker_NA), Runs1,0))]
temp1 <- temp1[, .(Runs1 = max(Runs1), Runs2 = max(Runs2)) , by = .(Batsman1, Batsman2)]
all_score <- match_data[,.(Ttl_runs = sum(runs_total)), by = .(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA))]
temp1 <- merge(temp1, all_score, all = TRUE)
temp1[, ':='(`Bman1%` = format(round((Runs1/Ttl_runs) * 100, 2), nsmall=2), `Bman2%` = format(round((Runs2/Ttl_runs) * 100, 2), nsmall=2))]
temp1[, Batsman := paste(Batsman1, Batsman2, sep = "/")]
setcolorder(temp1, c("Batsman1","Batsman2","Batsman","Runs1","Bman1%","Runs2","Bman2%","Ttl_runs"))

# High Score
high_score <- match_data[,.(score = sum(runs_total)), by = .(match_no, Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA))]
high_score <- high_score[, .(HS = max(score)), by = .(Batsman1, Batsman2)]

temp2 <- match_data[is.na(extras_wides), .(BF1 = .N), by = .(batsman_NA, non_striker_NA)]
temp2[, BF2 := 0]
temp2[, ':='(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), BF1 = ifelse(batsman_NA == pmin(batsman_NA, non_striker_NA), BF1,0),  BF2 = ifelse(batsman_NA == pmax(batsman_NA, non_striker_NA), BF1,0))]
temp2 <- temp2[, .(BF1 = max(BF1), `BF1%` = format(round((max(BF1)/sum(BF1,BF2)) * 100, 2), nsmall=2) ,BF2 = max(BF2), `BF2%` = format(round((max(BF2)/sum(BF1,BF2)) * 100, 2), nsmall=2)) , by = .(Batsman1, Batsman2)]
temp2 <- temp2[, Ttl_ball := BF1+BF2]

# Add dismissal data
temp3 <- match_data[!is.na(wicket_player_out), .(outs = .N), by = .(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), Pout = wicket_player_out)]
temp3 <- temp3[, .(Bman1 = pmin(Pout[1], Pout[2], na.rm = TRUE), Bman2 = pmax(Pout[1], Pout[2], na.rm = TRUE), Bman1_out = ifelse(pmin(Pout[1],Pout[2], na.rm = TRUE) == Pout[1], outs[1], outs[2]), Bman2_out = ifelse(pmax(Pout[1], Pout[2], na.rm = TRUE) == Pout[1], outs[1], outs[2])), by = .(Batsman1, Batsman2)]
temp3[Bman1 == Bman2, ':='(Bman1_out = ifelse(Batsman1 == Bman1, Bman1_out, NA), Bman2_out = ifelse(Batsman2 == Bman1, Bman2_out, NA))]
temp3[, c("Bman1","Bman2") := NULL]
temp3[, Ttl_out := rowSums(.SD ,na.rm = T), .SDcols = c("Bman1_out","Bman2_out")]

temp <- merge(temp, temp1, all = TRUE, by = c("Batsman1","Batsman2"))
temp <- merge(temp, high_score, all = TRUE)
temp <- merge(temp, temp2, all = TRUE)
temp <- merge(temp, temp3, all = TRUE)

temp[BF1 > 0, SR1 := format(round((Runs1/BF1) * 100, 2), nsmall = 2)]
temp[BF2 > 0, SR2 := format(round((Runs2/BF2) * 100, 2), nsmall = 2)]
temp[Ttl_ball > 0, Agg_SR := format(round((Ttl_runs/Ttl_ball) * 100, 2), nsmall = 2)]

temp[!is.na(Bman1_out), Avg1 := format(round(Runs1/Bman1_out, 2), nsmall = 2)]
temp[!is.na(Bman2_out), Avg2 := format(round(Runs2/Bman2_out, 2), nsmall = 2)]
temp[!is.na(Bman2_out), Agg_avg := format(round(Ttl_runs/Ttl_out, 2), nsmall = 2)]

temp1 <- match_data[runs_batsman == 4 & is.na(runs_non_boundary), .(`4s_1` = .N), by = .(batsman_NA, non_striker_NA)]
temp1[, `4s_2` := 0]
temp1[, ':='(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), `4s_1` = ifelse(batsman_NA == pmin(batsman_NA, non_striker_NA), `4s_1`,0),  `4s_2` = ifelse(batsman_NA == pmax(batsman_NA, non_striker_NA), `4s_1`,0))]
temp1 <- temp1[, .(`4s_1` = max(`4s_1`), `4s_2` = max(`4s_2`)) , by = .(Batsman1, Batsman2)]
temp1 <- temp1[, Ttl_4s := `4s_1`+`4s_2`]

temp2 <- match_data[runs_batsman == 6 & is.na(runs_non_boundary), .(`6s_1` = .N), by = .(batsman_NA, non_striker_NA)]
temp2[, `6s_2` := 0]
temp2[, ':='(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), `6s_1` = ifelse(batsman_NA == pmin(batsman_NA, non_striker_NA), `6s_1`,0),  `6s_2` = ifelse(batsman_NA == pmax(batsman_NA, non_striker_NA), `6s_1`,0))]
temp2 <- temp2[, .(`6s_1` = max(`6s_1`), `6s_2` = max(`6s_2`)) , by = .(Batsman1, Batsman2)]
temp2 <- temp2[, Ttl_6s := `6s_1`+`6s_2`]

temp <- merge(temp, temp1, all = TRUE)
temp <- merge(temp, temp2, all = TRUE)
temp[, c("Batsman1","Batsman2") := NULL]

modify_excel_sheet(sheetName = "OverallStats", wbPath = "./stats/Partnerships_Stats.xlsx", data = temp)

## Season Stats

match_played_data <- merge(match_data, info[,.(match_no, Season = year(date))], by = "match_no")
temp <- match_played_data[, .(Inn = length(unique(match_no))), by = .(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), Season)]

temp1 <- match_played_data[, .(Runs1 = sum(runs_batsman)), by = .(batsman_NA, non_striker_NA, Season)]
temp1[, Runs2 := 0]
temp1[, ':='(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), Runs1 = ifelse(batsman_NA == pmin(batsman_NA, non_striker_NA), Runs1,0),  Runs2 = ifelse(batsman_NA == pmax(batsman_NA, non_striker_NA), Runs1,0))]
temp1 <- temp1[, .(Runs1 = max(Runs1), Runs2 = max(Runs2)) , by = .(Batsman1, Batsman2, Season)]
all_score <- match_played_data[,.(Ttl_runs = sum(runs_total)), by = .(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), Season)]
temp1 <- merge(temp1, all_score, all = TRUE)
temp1[, ':='(`Bman1%` = format(round((Runs1/Ttl_runs) * 100, 2), nsmall=2), `Bman2%` = format(round((Runs2/Ttl_runs) * 100, 2), nsmall=2))]
temp1[, Batsman := paste(Batsman1, Batsman2, sep = "/")]
setcolorder(temp1, c("Batsman1","Batsman2","Batsman","Season","Runs1","Bman1%","Runs2","Bman2%","Ttl_runs"))

# High Score
high_score <- match_played_data[,.(score = sum(runs_total)), by = .(match_no, Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), Season)]
high_score <- high_score[, .(HS = max(score)), by = .(Batsman1, Batsman2, Season)]

temp2 <- match_played_data[is.na(extras_wides), .(BF1 = .N), by = .(batsman_NA, non_striker_NA, Season)]
temp2[, BF2 := 0]
temp2[, ':='(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), BF1 = ifelse(batsman_NA == pmin(batsman_NA, non_striker_NA), BF1,0),  BF2 = ifelse(batsman_NA == pmax(batsman_NA, non_striker_NA), BF1,0))]
temp2 <- temp2[, .(BF1 = max(BF1), `BF1%` = format(round((max(BF1)/sum(BF1,BF2)) * 100, 2), nsmall=2) ,BF2 = max(BF2), `BF2%` = format(round((max(BF2)/sum(BF1,BF2)) * 100, 2), nsmall=2)) , by = .(Batsman1, Batsman2, Season)]
temp2 <- temp2[, Ttl_ball := BF1+BF2]

# Add dismissal data
temp3 <- match_played_data[!is.na(wicket_player_out), .(outs = .N), by = .(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), Season, Pout = wicket_player_out)]
temp3 <- temp3[, .(Bman1 = pmin(Pout[1], Pout[2], na.rm = TRUE), Bman2 = pmax(Pout[1], Pout[2], na.rm = TRUE), Bman1_out = ifelse(pmin(Pout[1],Pout[2], na.rm = TRUE) == Pout[1], outs[1], outs[2]), Bman2_out = ifelse(pmax(Pout[1], Pout[2], na.rm = TRUE) == Pout[1], outs[1], outs[2])), by = .(Batsman1, Batsman2, Season)]
temp3[Bman1 == Bman2, ':='(Bman1_out = ifelse(Batsman1 == Bman1, Bman1_out, NA), Bman2_out = ifelse(Batsman2 == Bman1, Bman2_out, NA))]
temp3[, c("Bman1","Bman2") := NULL]
temp3[, Ttl_out := rowSums(.SD ,na.rm = T), .SDcols = c("Bman1_out","Bman2_out")]

temp <- merge(temp, temp1, all = TRUE, by = c("Batsman1","Batsman2","Season"))
temp <- merge(temp, high_score, all = TRUE)
temp <- merge(temp, temp2, all = TRUE)
temp <- merge(temp, temp3, all = TRUE)

temp[BF1 > 0, SR1 := format(round((Runs1/BF1) * 100, 2), nsmall = 2)]
temp[BF2 > 0, SR2 := format(round((Runs2/BF2) * 100, 2), nsmall = 2)]
temp[Ttl_ball > 0, Agg_SR := format(round((Ttl_runs/Ttl_ball) * 100, 2), nsmall = 2)]

temp[!is.na(Bman1_out), Avg1 := format(round(Runs1/Bman1_out, 2), nsmall = 2)]
temp[!is.na(Bman2_out), Avg2 := format(round(Runs2/Bman2_out, 2), nsmall = 2)]
temp[!is.na(Bman2_out), Agg_avg := format(round(Ttl_runs/Ttl_out, 2), nsmall = 2)]

temp1 <- match_played_data[runs_batsman == 4 & is.na(runs_non_boundary), .(`4s_1` = .N), by = .(batsman_NA, non_striker_NA, Season)]
temp1[, `4s_2` := 0]
temp1[, ':='(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), `4s_1` = ifelse(batsman_NA == pmin(batsman_NA, non_striker_NA), `4s_1`,0),  `4s_2` = ifelse(batsman_NA == pmax(batsman_NA, non_striker_NA), `4s_1`,0))]
temp1 <- temp1[, .(`4s_1` = max(`4s_1`), `4s_2` = max(`4s_2`)) , by = .(Batsman1, Batsman2, Season)]
temp1 <- temp1[, Ttl_4s := `4s_1`+`4s_2`]

temp2 <- match_played_data[runs_batsman == 6 & is.na(runs_non_boundary), .(`6s_1` = .N), by = .(batsman_NA, non_striker_NA, Season)]
temp2[, `6s_2` := 0]
temp2[, ':='(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), `6s_1` = ifelse(batsman_NA == pmin(batsman_NA, non_striker_NA), `6s_1`,0),  `6s_2` = ifelse(batsman_NA == pmax(batsman_NA, non_striker_NA), `6s_1`,0))]
temp2 <- temp2[, .(`6s_1` = max(`6s_1`), `6s_2` = max(`6s_2`)) , by = .(Batsman1, Batsman2, Season)]
temp2 <- temp2[, Ttl_6s := `6s_1`+`6s_2`]

temp <- merge(temp, temp1, all = TRUE)
temp <- merge(temp, temp2, all = TRUE)
temp[, c("Batsman1","Batsman2") := NULL]

rm(match_played_data)

modify_excel_sheet(sheetName = "SeasonStats", wbPath = "./stats/Partnerships_Stats.xlsx", data = temp)


## Against Teams

setnames(opposing_team, "team", "Team")

match_played_data <- merge(match_data, opposing_team, by.x = c("match_no","L2"), by.y = c("match_no","innings"))

temp <- match_played_data[, .(Inn = length(unique(match_no))), by = .(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), Team)]

temp1 <- match_played_data[, .(Runs1 = sum(runs_batsman)), by = .(batsman_NA, non_striker_NA, Team)]
temp1[, Runs2 := 0]
temp1[, ':='(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), Runs1 = ifelse(batsman_NA == pmin(batsman_NA, non_striker_NA), Runs1,0),  Runs2 = ifelse(batsman_NA == pmax(batsman_NA, non_striker_NA), Runs1,0))]
temp1 <- temp1[, .(Runs1 = max(Runs1), Runs2 = max(Runs2)) , by = .(Batsman1, Batsman2, Team)]
all_score <- match_played_data[,.(Ttl_runs = sum(runs_total)), by = .(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), Team)]
temp1 <- merge(temp1, all_score, all = TRUE)
temp1[, ':='(`Bman1%` = format(round((Runs1/Ttl_runs) * 100, 2), nsmall=2), `Bman2%` = format(round((Runs2/Ttl_runs) * 100, 2), nsmall=2))]
temp1[, Batsman := paste(Batsman1, Batsman2, sep = "/")]
setcolorder(temp1, c("Batsman1","Batsman2","Batsman","Team","Runs1","Bman1%","Runs2","Bman2%","Ttl_runs"))

# High Score
high_score <- match_played_data[,.(score = sum(runs_total)), by = .(match_no, Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), Team)]
high_score <- high_score[, .(HS = max(score)), by = .(Batsman1, Batsman2, Team)]

temp2 <- match_played_data[is.na(extras_wides), .(BF1 = .N), by = .(batsman_NA, non_striker_NA, Team)]
temp2[, BF2 := 0]
temp2[, ':='(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), BF1 = ifelse(batsman_NA == pmin(batsman_NA, non_striker_NA), BF1,0),  BF2 = ifelse(batsman_NA == pmax(batsman_NA, non_striker_NA), BF1,0))]
temp2 <- temp2[, .(BF1 = max(BF1), `BF1%` = format(round((max(BF1)/sum(BF1,BF2)) * 100, 2), nsmall=2) ,BF2 = max(BF2), `BF2%` = format(round((max(BF2)/sum(BF1,BF2)) * 100, 2), nsmall=2)) , by = .(Batsman1, Batsman2, Team)]
temp2 <- temp2[, Ttl_ball := BF1+BF2]

# Add dismissal data
temp3 <- match_played_data[!is.na(wicket_player_out), .(outs = .N), by = .(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), Team, Pout = wicket_player_out)]
temp3 <- temp3[, .(Bman1 = pmin(Pout[1], Pout[2], na.rm = TRUE), Bman2 = pmax(Pout[1], Pout[2], na.rm = TRUE), Bman1_out = ifelse(pmin(Pout[1],Pout[2], na.rm = TRUE) == Pout[1], outs[1], outs[2]), Bman2_out = ifelse(pmax(Pout[1], Pout[2], na.rm = TRUE) == Pout[1], outs[1], outs[2])), by = .(Batsman1, Batsman2, Team)]
temp3[Bman1 == Bman2, ':='(Bman1_out = ifelse(Batsman1 == Bman1, Bman1_out, NA), Bman2_out = ifelse(Batsman2 == Bman1, Bman2_out, NA))]
temp3[, c("Bman1","Bman2") := NULL]
temp3[, Ttl_out := rowSums(.SD ,na.rm = T), .SDcols = c("Bman1_out","Bman2_out")]

temp <- merge(temp, temp1, all = TRUE, by = c("Batsman1","Batsman2","Team"))
temp <- merge(temp, high_score, all = TRUE)
temp <- merge(temp, temp2, all = TRUE)
temp <- merge(temp, temp3, all = TRUE)

temp[BF1 > 0, SR1 := format(round((Runs1/BF1) * 100, 2), nsmall = 2)]
temp[BF2 > 0, SR2 := format(round((Runs2/BF2) * 100, 2), nsmall = 2)]
temp[Ttl_ball > 0, Agg_SR := format(round((Ttl_runs/Ttl_ball) * 100, 2), nsmall = 2)]

temp[!is.na(Bman1_out), Avg1 := format(round(Runs1/Bman1_out, 2), nsmall = 2)]
temp[!is.na(Bman2_out), Avg2 := format(round(Runs2/Bman2_out, 2), nsmall = 2)]
temp[!is.na(Bman2_out), Agg_avg := format(round(Ttl_runs/Ttl_out, 2), nsmall = 2)]

temp1 <- match_played_data[runs_batsman == 4 & is.na(runs_non_boundary), .(`4s_1` = .N), by = .(batsman_NA, non_striker_NA, Team)]
temp1[, `4s_2` := 0]
temp1[, ':='(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), `4s_1` = ifelse(batsman_NA == pmin(batsman_NA, non_striker_NA), `4s_1`,0),  `4s_2` = ifelse(batsman_NA == pmax(batsman_NA, non_striker_NA), `4s_1`,0))]
temp1 <- temp1[, .(`4s_1` = max(`4s_1`), `4s_2` = max(`4s_2`)) , by = .(Batsman1, Batsman2, Team)]
temp1 <- temp1[, Ttl_4s := `4s_1`+`4s_2`]

temp2 <- match_played_data[runs_batsman == 6 & is.na(runs_non_boundary), .(`6s_1` = .N), by = .(batsman_NA, non_striker_NA, Team)]
temp2[, `6s_2` := 0]
temp2[, ':='(Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), `6s_1` = ifelse(batsman_NA == pmin(batsman_NA, non_striker_NA), `6s_1`,0),  `6s_2` = ifelse(batsman_NA == pmax(batsman_NA, non_striker_NA), `6s_1`,0))]
temp2 <- temp2[, .(`6s_1` = max(`6s_1`), `6s_2` = max(`6s_2`)) , by = .(Batsman1, Batsman2, Team)]
temp2 <- temp2[, Ttl_6s := `6s_1`+`6s_2`]

temp <- merge(temp, temp1, all = TRUE)
temp <- merge(temp, temp2, all = TRUE)
temp[, c("Batsman1","Batsman2") := NULL]

rm(match_played_data)
setnames(opposing_team, "Team", "team")

modify_excel_sheet(sheetName = "AgainstTeam", wbPath = "./stats/Partnerships_Stats.xlsx", data = temp)

# Wicketwise Partnership by Teams

setnames(opposing_team, "team", "opp_team")
match_played_data <- merge(match_data, team_innings, by.x = c("match_no","L2"), by.y = c("match_no","innings"))
match_played_data <- merge(match_played_data, opposing_team, by.x = c("match_no","L2"), by.y = c("match_no","innings"))

match_played_data[,Wicket := cumsum(!is.na(wicket_player_out)), by = .(match_no, L2)]
match_played_data[, Wicket := Wicket + 1]
match_played_data[, Wicket := shift(Wicket), by = .(match_no,L2)]
match_played_data[is.na(Wicket), Wicket := 1]

temp <- match_played_data[, .(opp_team = min(opp_team), score = sum(runs_total)), by = .(match_no, Batsman1 = pmin(batsman_NA, non_striker_NA), Batsman2 = pmax(batsman_NA, non_striker_NA), Wicket, team)]
temp <- temp[, .(match_no = match_no[which.max(score)], Batsman1 = Batsman1[which.max(score)], Batsman2 = Batsman2[which.max(score)], HS = max(score), opp_team = opp_team[which.max(score)]), by = .(Wicket, team)]
temp <- temp[order(team, Wicket)]
rm(match_played_data)

setnames(opposing_team, "opp_team", "team")

modify_excel_sheet(sheetName = "Pship_wkt_team", wbPath = "./stats/Partnerships_Stats.xlsx", data = temp)

### Common and Field Stats

## Batsman vs bowler stats

temp <- data.table()

innings <- match_data[, .(Innings = length(unique(match_no))), by = .(Batsman = batsman_NA, Bowler = bowler_NA)]


wickets <- match_data[wicket_kind %in% bowler_dismissal, .(Wkts = .N), by = .(Batsman = batsman_NA, Bowler = bowler_NA)]

runs <- match_data[, .(Runs = sum(runs_batsman)), by = .(Batsman = batsman_NA, Bowler = bowler_NA)]

balls <- match_data[is.na(extras_wides), .(BF = .N), by = .(Batsman = batsman_NA, Bowler = bowler_NA)]

no_balls <- match_data[!is.na(extras_noballs), .(NoBl = .N), by = .(Batsman = batsman_NA, Bowler = bowler_NA)]

temp <- merge(innings, wickets, all = TRUE)
temp <- merge(temp, runs, all = TRUE)
temp <- merge(temp, balls, all = TRUE)
temp <- merge(temp, no_balls, all = TRUE)

temp[, Avg := ifelse(is.na(Runs/Wkts), NA, format(round(Runs/Wkts,2),nsmall = 2))]
temp[, SR := ifelse(is.na(Runs/BF), NA, format(round(Runs/BF * 100,2),nsmall = 2))]

fours <- match_data[runs_batsman == 4 & is.na(runs_non_boundary), .('4s' = .N), by = .(Batsman = batsman_NA, Bowler = bowler_NA)]

sixes <- match_data[runs_batsman == 6 & is.na(runs_non_boundary), .('6s' = .N), by = .(Batsman = batsman_NA, Bowler = bowler_NA)]

temp <- merge(temp, fours, all = TRUE)
temp <- merge(temp, sixes, all = TRUE)

dots <- match_data[runs_batsman == 0, .(Dots = .N), by = .(Batsman = batsman_NA, Bowler = bowler_NA)]
temp <- merge(temp, dots, all = TRUE)

temp[, Bowl_SR := ifelse(is.na(rowSums(.SD, na.rm = T)/Wkts), NA, format(round(rowSums(.SD, , na.rm = T)/Wkts,2),nsmall = 2)), .SDcols = c("BF","NoBl")]

temp[, NoBl := NULL]

rm(list = c("innings","wickets","runs","balls","no_balls","fours","sixes","dots"))

modify_excel_sheet(sheetName = "BowlerVsBatsman", wbPath = "./stats/Common_and_Field_Stats.xlsx", data = temp)

## MOM total

temp <- info[, .(matches = .N), by = MOM][order(-matches)]

modify_excel_sheet(sheetName = "MOM", wbPath = "./stats/Common_and_Field_Stats.xlsx", data = temp)

## Umpires and matches

temp <- data.table(info[, table(c(umpire1, umpire2))])
setnames(temp, c("V1","N"), c("Umpire","Matches"))
temp <- temp[order(-Matches)]

modify_excel_sheet(sheetName = "Umpire_matches", wbPath = "./stats/Common_and_Field_Stats.xlsx", data = temp)

## Export Info

temp <- copy(info)
temp[, ':='(Umpires = paste(umpire1, umpire2, sep = ", "), Teams = paste(team1, team2, sep = ", "))]
modify_excel_sheet(sheetName = "allMatchInfo", wbPath = "./stats/Common_and_Field_Stats.xlsx", data = temp)
