library(readxl)

# Step 1: Set folder path. (The folder you store all the excel files)
# One thing important: if the folder path showed this pattern: 
#   C:\Users\Username\Desktop
# Have to change into this pattern:
#   C:\\Users\\Username\\Desktop
# Otherwise R cannot read it!
folder_path <- "your\\own\\folder\\path"

# set all the excel files path
# Do not need to change anything here.
files <- list.files(path = folder_path, pattern = "\\.xlsx$", full.names = TRUE)

# Create a new data frame to store the excels that missing columns we want
missing_columns_info <- data.frame(File = character(), Sheet = character(), stringsAsFactors = FALSE)

# Step 2: Set the column names that you want to find in those excels
# In here I want to find columns that have "code" or "codes"
# Remember! to keep them all lower case! No "Code" or "Codes"
required_columns <- c("code", "codes")

# Process all the files in the folder that we set
for (file in files) {
  sheets <- excel_sheets(file)  # get all the sheet names
  
  for (sheet in sheets) {
    df <- read_excel(file, sheet = sheet)  # read all the sheets
    column_names <- tolower(names(df))  # turn all the columns in excels to lower case 
                                        # To match our required_columns
    
    # Checking for matching column names
    if (!any(column_names %in% required_columns)) {
      missing_columns_info <- rbind(
        missing_columns_info,
        data.frame(File = basename(file), Sheet = sheet, stringsAsFactors = FALSE)
      )
    }
  }
}

# Checking for excels that does not have required column names
print(missing_columns_info)

# Returning results
if (nrow(missing_columns_info) == 0) {
  print("All the excel sheets at least includes one of the required column name: ")
} else {
  print("The excels list below is lacking required column name：")
  print(missing_columns_info)
}
