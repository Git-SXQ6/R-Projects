library(readxl)
library(dplyr)
library(tidyr)
library(openxlsx)

setwd("C:\\Users\\25022\\Desktop")

# Read Excel file without headers
df <- read_excel("Parent and child coding.xlsx", col_names = FALSE) %>%
  as.data.frame()

# If no errors, process multi-row headers
colnames(df) <- paste(df[1, ], df[2, ], sep = "_")
df <- df[-c(1,2), ]  # Remove header rows

setwd("Coding")

raw_data <- list.files(, pattern = "\\.xlsx$", full.names = FALSE)

df_list <- lapply (raw_data, read_excel)

names(df_list) <- gsub("\\.xlsx$", "", raw_data)

label_counts <- list()

for (table in names(df_list)) {
  column_names <- tolower(names(df_list[[table]]))
  current_table <- df_list[[table]] |>
    select(contains("cod"))
  unique_labels <- unique(unlist(current_table))
  label_count <- table(unlist(current_table))
  label_count_df <- as.data.frame(label_count)
  colnames(label_count_df) <- c("Code", "Count")
  label_counts[[table]] <- label_count_df
}

all_labels <- unique(unlist(lapply(label_counts, `[[`, "Code")))

for (table_name in names(label_counts)) {
  current_table <- label_counts[[table_name]]
  missing_labels <- setdiff(all_labels, current_table$Code)
  # Check if there are missing labels
  if (length(missing_labels) > 0) {
    missing_data <- data.frame(Code = missing_labels, Count = 0)
    label_counts[[table_name]] <- bind_rows(current_table, missing_data) |> arrange(Code)
  } else {
    # If no missing labels, just keep the current table as is
    label_counts[[table_name]] <- current_table
  }
 
  label_counts[[table_name]] <- bind_rows(current_table, missing_data) |>
    arrange(Code)
}

# 5. Combine all label count tables into one
final_df <- bind_rows(label_counts, .id = "Table")

final_df <- final_df |>
  pivot_wider(
    id_cols = Table,
    names_from = Code, 
    values_from = Count,
    values_fn = sum
  )

write.xlsx(final_df, "T3Manchester.xlsx")

