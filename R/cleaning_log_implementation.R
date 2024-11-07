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

df_remove_survey_cl <- df_filled_cl %>% 
  filter(change_type %in% c("remove_survey"))