library(tidyverse)
library(readxl)
library(openxlsx)
## Create a variable for the path to the excel file
path_ex <- "/Users/oliverschramm/Documents/Documents - MacBook Pro (2)/GRADUATE SCHOOL/DSCC Projects/CCJ Bail Reform/aoicprobationfiles/aoicprobationfiles/2024 circuit datafiles/2024_23rd Circuit Data.xlsx"#str_c("aoicprobationfiles/aoicprobationfiles/", folder, "/", file)
    
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
