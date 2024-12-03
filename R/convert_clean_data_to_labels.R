library(tidyverse)
library(supporteR)

# clean data
clean_loc_data <- "inputs/diagnostic_of_informality_cleaned_data.xlsx"

# main data
clean_data_nms <- names(readxl::read_excel(path = clean_loc_data, n_max = 1000, sheet = "cleaned_data"))
clean_c_types <- ifelse(str_detect(string = clean_data_nms, pattern = "_other$"), "text", "guess")
df_main_clean_data <- readxl::read_excel(clean_loc_data, col_types = clean_c_types, sheet = "cleaned_data")

# loops
# roster
clean_data_nms_r_roster <- names(readxl::read_excel(path = clean_loc_data, n_max = 3000, sheet = "cleaned_roster"))
clean_c_types_r_roster <- ifelse(str_detect(string = clean_data_nms_r_roster, pattern = "_other$"), "text", "guess")
df_clean_loop_r_roster <- readxl::read_excel(clean_loc_data, col_types = clean_c_types_r_roster, sheet = "cleaned_roster")

# tool
loc_tool <- "inputs/Diagnostic_of_informality_Tool.xlsx"
df_survey <- readxl::read_excel(loc_tool, sheet = "survey")
df_choices <- readxl::read_excel(loc_tool, sheet = "choices")