library(tidyverse)
library(readxl)
library(openxlsx)
## get the folders we are interested in from aoicprobationfiles
folders <- list.files("aoicprobationfiles/aoicprobationfiles")[str_detect(list.files("aoicprobationfiles/aoicprobationfiles"), "datafiles")]
## then get each file that we want in the folder
files <- map(str_c("aoicprobationfiles/aoicprobationfiles/", folders), list.files) ## Map iterates the list.files function over each folder in 'folders'
## name the list for for each folder
names(files) <- folders

for (i in 1:length(folders)) {
  folder <- folders[i]
  ## Separating the year from the Circuits. Files[[1]] is 2022, files[[2]] is 2023, etc.
  for (j in 1:length(files[[i]])) {
    file <- files[[i]][j]
    map(str_split(files[[i]], pattern = "_"), ~ .[1]) %>%  unlist()
    map(str_split(files[[i]], pattern = "_"), ~ .[2]) %>%  unlist()
    ## Removes Data.xlsx from each circuits data file (Gets district)
    str_remove(map(str_split(files[[i]], pattern = "_"), ~ .[2]) %>%  unlist(), " Data.xlsx")
    ## Files is a list of different sizes. We want to go and access the different sizes for years.
    map(1:length(files), ~map(str_split(files[[.x]], pattern = "_"), ~ .[1]) %>% unlist())
    ## Access different sizes for circuits/districts
    map(1:length(files), ~map(str_split(files[[.x]], pattern = "_"), ~ .[2]) %>% unlist())
    ## Turn the years into a vector
    years <- map(1:length(files), ~map(str_split(files[[.x]], pattern = "_"), ~ .[1]) %>% unlist())
    ## Turn the courts into a vector
    courts <- map(1:length(files), ~str_remove(map(str_split(files[[1]], pattern = "_"), ~ .[2]) %>%  unlist(), " Data.xlsx"))

    ## Create a variable for the path to the excel file
    path_ex <- str_c("aoicprobationfiles/aoicprobationfiles/", folder, "/", file)
    
    ## This reads all of the sheets except for "Table of Contents"
    sheets <- excel_sheets(path_ex)[-1]
    
    ## Looping through the sheets
    for (sheet in sheets){
      read <- read_excel(path = path_ex, sheet = sheet, skip = 1)
      ## Creating the transpose of the read file
      dat <- t(read)
      names(dat) <- t(read)
      
      df <- as.data.frame(dat) %>% 
        rownames_to_column("month")
      
      names(df) <- df[1,]
      df <- df[-1,]
      
      ## Checking for and removing any column that is completely NA
      na_cols <- is.na(names(df))
      df_clean <- df[, !na_cols, drop = FALSE]
      ## Making everything numeric except for columns that have character values 
      ## and selecting just the columns for each month and gearing up to write .xlsx file
      ### find those columns that are not missing or contain a string of some kind (not numbers)
      strings <- sapply(df_clean[1,], function(x) !is.na(x) && !grepl("^[0-9.]+$", x))
      ## those columns that are strings
      df_clean[strings == TRUE]
      ### those that are numeric make numeric
      df_clean[strings == FALSE] = map(df_clean[strings == FALSE], as.numeric)
      
      lapply(df_clean[14:28, 1:length(names(df_clean))], function(x) sum(is.na(x)))
      
      df_clean[14:28, 1:length(names(df_clean))] %>% view()
      ## all of these are percentages of data we already keep
      ## can remove
      
      ready_file <- df_clean[1:12, 1:length(names(df_clean))] %>% 
        rename(month = ...1) %>% 
        mutate(circuit_year = sheet, .before = month)
      
      ## Writing out to excel
      write.xlsx(ready_file, paste0("transposed_files/", str_replace_all(sheet," ", ""), ".xlsx"))
    }
  }
}
