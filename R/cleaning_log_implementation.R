library(tidyverse)
library(glue)
library(cleaningtools)
library(httr)
library(supporteR)

loc_data <- "inputs/Diagnostic_of_informality_Data.xlsx"

# main data
data_nms <- names(readxl::read_excel(path = loc_data, n_max = 300))
c_types <- ifelse(str_detect(string = data_nms, pattern = "_other$"), "text", "guess")
df_tool_data <- readxl::read_excel(loc_data, col_types = c_types)

# loops
# roster
data_nms_r_roster <- names(readxl::read_excel(path = loc_data, n_max = 300, sheet = "hh_roster"))
c_types_r_roster <- ifelse(str_detect(string = data_nms_r_roster, pattern = "_other$"), "text", "guess")
df_loop_r_roster <- readxl::read_excel(loc_data, col_types = c_types_r_roster, sheet = "hh_roster")

# tool
loc_tool <- "inputs/Diagnostic_of_informality_Tool.xlsx"
df_survey <- readxl::read_excel(loc_tool, sheet = "survey")
df_choices <- readxl::read_excel(loc_tool, sheet = "choices")

df_filled_cl <- readxl::read_excel("inputs/combined_checks_nam_diagnostic.xlsx", sheet = "cleaning_log") %>% 
  filter(!is.na(reviewed), !question %in% c("_index"), !uuid %in% c("all"))

# surveys for deletion
df_remove_survey_cl <- df_filled_cl %>% 
  filter(change_type %in% c("remove_survey"))

# check pii ---------------------------------------------------------------
pii_from_data <- cleaningtools::check_pii(dataset = df_tool_data, element_name = "checked_dataset", uuid_column = "_uuid")
pii_from_data$potential_PII

# then determine wich columns to remove from both the raw and clean data
cols_to_remove <- c("audit", "audit_URL", 
                    "latitude", "longitude", "geopoint",
                    "instance_name", "_geopoint_latitude", "_geopoint_longitude",
                    "_geopoint_altitude", "_geopoint_precision")

# Main dataset ------------------------------------------------------------

# filtered log
df_filled_cl_main <- df_filled_cl %>% 
  filter(is.na(sheet))

# updating the main dataset with new columns

df_data_with_added_cols <- cts_add_new_sm_choices_to_data(input_df_tool_data = df_tool_data,
                                                          input_df_filled_cl = df_filled_cl_main, 
                                                          input_df_survey = df_survey,
                                                          input_df_choices = df_choices)
