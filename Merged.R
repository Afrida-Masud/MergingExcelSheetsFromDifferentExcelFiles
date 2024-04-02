# Install the packages
install.packages("readxl")
install.packages("dplyr")
install.packages("openxlsx")
# Loading the Packages
library(openxlsx)
library(readxl)
library(dplyr)
#Setting Working Directory
setwd("~/Capstone Projects/onet/db_28_2_excel/db_28_2_excel")
# Link to download the O*net zip file: https://www.onetcenter.org/database.html#all-files 
# It is under the all file
# Download the data in your working directory and unzip it

# Path to the folder containing all Excel files
folder_path <- "C:/Users/amasud/OneDrive - IL State University/Capstone Projects/onet/db_28_2_excel/db_28_2_excel"

# List all Excel files in the folder
excel_files <- list.files(path = folder_path, pattern = "\\.xlsx$", full.names = TRUE)

# Create a new workbook
merged_workbook <- createWorkbook()

# Iterate over each Excel file
for (file in excel_files) {
  # Read all sheets from the Excel file
  sheets <- getSheetNames(file)
  
  # Iterate over each sheet
  for (sheet in sheets) {
    # Read the data from the sheet
    data <- read.xlsx(file, sheet)
    
    # Add the data to the merged workbook as a new sheet
    addWorksheet(merged_workbook, sheet)
    writeData(merged_workbook, sheet, data)
  }
}

# Save the merged workbook
saveWorkbook(merged_workbook, "merged_output.xlsx")

