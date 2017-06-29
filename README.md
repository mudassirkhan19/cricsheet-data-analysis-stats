# Read Before Using
This is an R project that downloads, reads and aggregates data from cricsheet (T20 and ODIs) and provides a few stats in form or excel spreadsheets. It basically performs four tasks.

### 1. Download Data :

Cricsheet data updates after every few days and it is cumbersome to download manually all put it in the required folder. downloadData.R downloads the updated zip file, compares it with the existing zip files and extracts only the new files added. If you are using more than one zip files and aggregating them make sure to change the:
a) Link of the zip file to download and 
b) Path of the new zip file
c) Path of the old zip file

### 2. Convert to CSV :

After downloading and extracting the yaml files downloadData.R converts the yaml data to CSV and puts each match data into two CSV files. If you want to do this manually just use the convert_cricsheet_csv.R function.

### 3. Load Data into R environment :

loadData.R reads all files from match_data folder into the workspace. Edit it to read only the files you require. The scripts reads IPL, T20 and BBL data.

All the above three tasks can be performed by using prepare.R.

### 4. Print Stats :

stats.R generates some predefined stats into the stats folder.

## Scraped Data
If you want to add some scraped data using the [scraper](https://github.com/mudassirkhan19/cricinfo-ballwise-scraper), put the match_data and info files scraped using it in match_data_converted folder and run cleanScrapedData.R script, then move the modified files to match_data folder. Then run the above scripts and it should work smoothly.
