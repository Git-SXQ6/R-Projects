library(readxl)
library(dplyr)
library(tidyr)

# Set the folder containing Excel files
folder_path <- "C:\\Users\\25022\\Desktop\\Coding"

# Get all .xlsx file paths
files <- list.files(path = folder_path, pattern = "\\.xlsx$", full.names = TRUE)

# Define possible column names (case-insensitive)
valid_columns <- tolower(c("code", "codes"))

# Initialize an empty list to store results
all_results <- list()
all_codes <- c()  # To store all unique codes found

# Step 1: Extract codes from all files and sheets
for (file in files) {
  sheets <- excel_sheets(file)  # Get all sheet names
  
  for (sheet in sheets) {
    df <- read_excel(file, sheet = sheet)  # Read sheet
    col_names <- tolower(names(df))  # Convert column names to lowercase
    
    # Find the first matching column name
    matched_col <- names(df)[which(col_names %in% valid_columns)][1]
    
    if (!is.null(matched_col)) {
      # Select the column and remove NA values
      code_data <- df[[matched_col]] %>% na.omit()
      
      # Count occurrences of each unique value
      code_summary <- as.data.frame(table(code_data))
      
      # Append to results list
      all_results <- append(all_results, list(data.frame(File = basename(file), 
                                                         Sheet = sheet, 
                                                         Code = code_summary$code_data, 
                                                         Count = code_summary$Freq, 
                                                         stringsAsFactors = FALSE)))
      
      # Collect all unique codes found
      all_codes <- unique(c(all_codes, as.character(code_summary$code_data)))
    }
  }
}

# Step 2: Create a complete table with all unique codes
if (length(all_results) > 0) {
  combined_results <- bind_rows(all_results)
  
  # Ensure every file has an entry for each unique code
  complete_results <- expand.grid(File = unique(combined_results$File), 
                                  Code = all_codes, 
                                  stringsAsFactors = FALSE) %>%
    left_join(combined_results, by = c("File", "Code")) %>%
    mutate(Count = ifelse(is.na(Count), 0, Count))  # Fill missing counts with 0
  
  # Pivot the data to wide format (files as rows, codes as columns)
  wide_results <- complete_results %>%
    spread(key = Code, value = Count, fill = 0)
} else {
  wide_results <- data.frame(File = character(), stringsAsFactors = FALSE)
}

# Print results
print(wide_results)

# Name your data summary in a CSV file
write.csv(wide_results, "Name_Your_File_Here.csv", row.names = FALSE)

if (nrow(wide_results) == 0) {
  print("No coding data found in any of the files.")
} else {
  print("Code counting completed.")
}
