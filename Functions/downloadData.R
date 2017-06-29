downloadData <- function(){
    require(data.table)
    dest <- "./match_data/cricsheet"
    
    ### IPL
    ## Update
    link_Ipl <- "http://cricsheet.org/downloads/ipl.zip"
    zipIpl <- "./match_data/cricsheet/ipl.zip"
    zipOld_Ipl <- "./match_data/cricsheet/ipl_curr.zip"
    download.file(link_Ipl, zipIpl)
    
    ## check for files that don't exist in the current zip file
    temp <- unzip(zipIpl, list = TRUE)$Name
    if(file.exists(zipOld_Ipl)){
        temp1 <- unzip(zipOld_Ipl, list = TRUE)$Name
        temp2 <- setdiff(temp,temp1)
    }
    else{
        temp2 <- temp
    }
    
    # If there is a new file extract it
    if(length(temp2)){
        unzip(zipIpl, files = temp2, exdir = dest)
        source('./Functions/convert_cricsheet_csv.R')
        convertCricsheetData(source = dest, destination = "./match_data/IPL/")
        file.remove(list.files(path = dest, pattern = ".yaml$", full.names = TRUE))
    }
    
    if(file.exists(zipOld_Ipl)){file.remove(zipOld_Ipl)}
    file.rename(zipIpl, zipOld_Ipl)
    
    ### T20
    ## Update
    link_T20 <- "http://cricsheet.org/downloads/t20s_male.zip"
    zipT20 <- "./match_data/cricsheet/T20.zip"
    zipOld_T20 <- "./match_data/cricsheet/T20_curr.zip"
    download.file(link_T20, zipT20)
    
    ## check for files that don't exist in the current zip file
    temp <- unzip(zipT20, list = TRUE)$Name
    if(file.exists(zipOld_T20)){
        temp1 <- unzip(zipOld_T20, list = TRUE)$Name
        temp2 <- setdiff(temp,temp1)
    }
    else{
        temp2 <- temp
    }
    
    # If there is a new file extract it
    if(length(temp2)){
        unzip(zipT20, files = temp2, exdir = dest)
        source('./Functions/convert_cricsheet_csv.R')
        convertCricsheetData(source = dest, destination = "./match_data/T20/")
        file.remove(list.files(path = dest, pattern = ".yaml$", full.names = TRUE))
    }
    
    if(file.exists(zipOld_T20)){file.remove(zipOld_T20)}
    file.rename(zipT20, zipOld_T20)
    
    ### BBL
    ## Update
    link_Bbl <- "http://cricsheet.org/downloads/bbl.zip"
    zipBbl <- "./match_data/cricsheet/bbl.zip"
    zipOld_Bbl <- "./match_data/cricsheet/bbl_curr.zip"
    download.file(link_Bbl, zipBbl)
    
    ## check for files that don't exist in the current zip file
    temp <- unzip(zipBbl, list = TRUE)$Name
    if(file.exists(zipOld_Bbl)){
        temp1 <- unzip(zipOld_Bbl, list = TRUE)$Name
        temp2 <- setdiff(temp,temp1)
    }
    else{
        temp2 <- temp
    }
    
    # If there is a new file extract it
    if(length(temp2)){
        unzip(zipBbl, files = temp2, exdir = dest)
        source('./Functions/convert_cricsheet_csv.R')
        convertCricsheetData(source = dest, destination = "./match_data/BBL/")
        file.remove(list.files(path = dest, pattern = ".yaml$", full.names = TRUE))
    }
    
    if(file.exists(zipOld_Bbl)){file.remove(zipOld_Bbl)}
    file.rename(zipBbl, zipOld_Bbl)
    
}