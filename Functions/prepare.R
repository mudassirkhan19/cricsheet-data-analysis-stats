source('./Functions/downloadData.R')
downloadData()

source('./Functions/loadData.R')
req_data <- loadData()
match_data <- req_data[[1]]
info <- req_data[[2]]
team_innings <- req_data[[3]]
opposing_team <- req_data[[4]]
bowler_dismissal <- req_data[[5]]
rm(req_data)

# Add excel modifying function
source('./Functions/modify_excel_sheet.R')

