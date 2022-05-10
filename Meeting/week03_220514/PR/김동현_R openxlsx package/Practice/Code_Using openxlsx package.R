### Code : Using openxlsx package
### Writer : Donghyeon Kim
### Announcement Date : 2022.05.14

##########################################

xlsx_function <- function(file_x) {
  
  ## 1. Read xlsx file
  library(readxl) # package for read_xlsx function
  rawdata <- read_xlsx(file_x, sheet = 1)
  
  ## 2. Save as xlsx file by spliting data by energy sources
  energy <- unique(rawdata$에너지원)
  
    # 2-1) Create the first xlsx file while saving the first energy source
    library(tidyverse)
    save_data <- rawdata %>% filter(., `에너지원` == energy[1]) # First energy source : the light of the sun(태양광)
    
    library(openxlsx)
    if(file.exists("Result.xlsx") == FALSE) {
      write.xlsx(save_data, file = "Result.xlsx", sheetName = energy[1])
    }
    rm(save_data)
    
    # 2-2) Save different energy sources on each sheet in an Excel file created earlier
    for(i in 2:length(energy)) {
      save_data <- rawdata %>% filter(., `에너지원` == energy[i]) # Energy source from second to last
      
      workbook <- loadWorkbook("Result.xlsx") # Load workbook
      addWorksheet(workbook, sheetName = energy[i]) # Add sheet on workbook
      writeData(workbook, sheet = energy[i], save_data) # Write data on added sheet of workbook
      saveWorkbook(workbook, "Result.xlsx", overwrite = TRUE)
      rm(list = c("workbook", "save_data"))
    }
}

##########################################

## Running a User-defined function ##

# 1. Set working directory
dr <- "C:/"
foldname <- paste0(dr, "Users/mazy4/Desktop")
setwd(foldname)
getwd() # Check working directory

# 2. Set the file name to load
filename <- "Rawdata example.xlsx"

# 3. Execution
xlsx_function(filename)
