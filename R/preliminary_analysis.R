library(tidyverse)
library(srvyr)
library(supporteR)
library(analysistools)

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

# loa // list of analysis
all_loa <- read_csv("inputs/r_loa_nam_diagnostic.csv")


# data with composites ----------------------------------------------------

# main data with composites
df_data_with_composites <- df_main_clean_data %>% 
  mutate(i.respondent_age = ifelse(is.na(next_resp_age), init_resp_age, next_resp_age),
         i.respondent_gender = ifelse(is.na(next_resp_gender), init_resp_gender, next_resp_gender))

# roster
df_clean_loop_r_roster_with_composites <- df_clean_loop_r_roster

# analysis - main -------------------------------------------------

# main
df_main <- df_data_with_composites
# survey object
main_svy <- as_survey(.data = df_main)

# loa
df_main_loa <- all_loa %>% 
  filter(dataset %in% c("main_data"))

# analysis
df_main_analysis <- analysistools::create_analysis(design = main_svy, 
                                                   loa = df_main_loa,
                                                   sm_separator = "/")


# analysis - roster -----------------------------------------------

# roster
df_roster <- df_clean_loop_r_roster_with_composites

# survey object
roster_svy <- as_survey(.data = df_roster)

# loa roster
df_roster_loa <- all_loa %>% 
  filter(dataset %in% c("roster"))

# analysis
df_roster_analysis <- analysistools::create_analysis(design = roster_svy, 
                                                     loa = df_roster_loa,
                                                     sm_separator = "/")


# analysis tables ---------------------------------------------------------

# combine the tables

df_combined_tables <- bind_rows(df_main_analysis$results_table %>% mutate(dataset = "main"),
                                df_roster_analysis$results_table %>% mutate(dataset = "roster")) %>% 
  filter(!(analysis_type %in% c("prop_select_one", "prop_select_multiple") & (is.na(analysis_var_value) | analysis_var_value %in% c("NA"))))

# df_analysis_table <- presentresults::create_table_variable_x_group(results_table = df_combined_tables)
# 
# presentresults::create_xlsx_variable_x_group(table_group_x_variable = df_analysis_table,
#                                              file_path = paste0("outputs/", butteR::date_file_prefix(), "_analysis_tables_nam_diagnostic.xlsx"),
#                                              table_sheet_name = "diagnostic",
#                                              overwrite = TRUE
# 
# )

# openxlsx::write.xlsx(df_combined_tables, 
#                      paste0("outputs/", butteR::date_file_prefix(),"_non_formatted_analysis_nam_diagnostic.xlsx"), overwrite = TRUE)

write_csv(df_combined_tables, paste0("outputs/", butteR::date_file_prefix(), "_non_formatted_analysis_nam_diagnostic.csv"), na="")
write_csv(df_combined_tables, paste0("outputs/non_formatted_analysis_nam_diagnostic.csv"), na="")
