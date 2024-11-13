library(tidyverse)
library(glue)
library(cleaningtools)
library(httr)
library(supporteR)

loc_data <- "inputs/Diagnostic_of_informality_Data.xlsx"

# main data
data_nms <- names(readxl::read_excel(path = loc_data, n_max = 300))
c_types <- ifelse(str_detect(string = data_nms, pattern = "_other$"), "text", "guess")
df_tool_data <- readxl::read_excel(loc_data, col_types = c_types) %>% 
  mutate(point_number = paste0("pt_", row_number()),
         location = case_when(interview_loc_level %in% c("municipality") ~ municipality,
                              interview_loc_level %in% c("town_council") ~ town_council,
                              interview_loc_level %in% c("village_council") ~ village_council,
                              interview_loc_level %in% c("settlement") ~ settlement))

# loops
# roster
data_nms_r_roster <- names(readxl::read_excel(path = loc_data, n_max = 300, sheet = "hh_roster"))
c_types_r_roster <- ifelse(str_detect(string = data_nms_r_roster, pattern = "_other$"), "text", "guess")
df_loop_r_roster <- readxl::read_excel(loc_data, col_types = c_types_r_roster, sheet = "hh_roster")

# tool
loc_tool <- "inputs/Diagnostic_of_informality_Tool.xlsx"
df_survey <- readxl::read_excel(loc_tool, sheet = "survey")
df_choices <- readxl::read_excel(loc_tool, sheet = "choices")

# cleaning log 
log_data <- "inputs/combined_checks_nam_diagnostic.xlsx"
log_nms <- names(readxl::read_excel(path = log_data, sheet = "cleaning_log", n_max = 6000))
c_types_log <- case_when(str_detect(string = log_nms, pattern = "index|reviewed") ~ "numeric", 
                         str_detect(string = log_nms, pattern = "sheet|other_text") ~ "text",
                         TRUE ~ "guess")
df_filled_cl <- readxl::read_excel(log_data, sheet = "cleaning_log", col_types = c_types_log) %>% 
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

# check the cleaning log
df_cl_review <- cleaningtools::review_cleaning_log(
  raw_dataset = df_data_with_added_cols,
  raw_data_uuid_column = "_uuid",
  cleaning_log = df_filled_cl_main,
  cleaning_log_change_type_column = "change_type",
  change_response_value = "change_response",
  cleaning_log_question_column = "question",
  cleaning_log_uuid_column = "uuid",
  cleaning_log_new_value_column = "new_value")

# filter log for cleaning
df_final_cleaning_log <- df_filled_cl_main %>% 
  filter(!question %in% c("duration_audit_sum_all_ms", "duration_audit_sum_all_minutes", "phone_consent"), 
         !uuid %in% c("all")) %>% 
  filter(!str_detect(string = question, pattern = "\\w+\\/$"))

# create the clean data from the raw data and cleaning log
df_cleaning_step <- cleaningtools::create_clean_data(
  raw_dataset = df_data_with_added_cols %>% select(-any_of(cols_to_remove)),
  raw_data_uuid_column = "_uuid",
  cleaning_log = df_final_cleaning_log,
  cleaning_log_change_type_column = "change_type",
  change_response_value = "change_response",
  NA_response_value = "blank_response",
  no_change_value = "no_action",
  remove_survey_value = "remove_survey",
  cleaning_log_question_column = "question",
  cleaning_log_uuid_column = "uuid",
  cleaning_log_new_value_column = "new_value")

# handle parent question columns
df_updating_sm_parents <- cts_update_sm_parent_cols(input_df_cleaning_step_data = df_cleaning_step, 
                                                    input_uuid_col = "_uuid",
                                                    input_point_id_col = "point_number",
                                                    input_collected_date_col = "today",
                                                    input_location_col = "location")

# tool data to support loops ----------------------------------------------

df_tool_support_data_for_loops <- df_updating_sm_parents$updated_sm_parents %>% 
  filter(!`_uuid` %in% df_remove_survey_cl$uuid) %>% 
  select(`_uuid`, today, enumerator_id, point_number, location)

# roster cleaning ---------------------------------------------------------

# filtered log
df_filled_cl_roster <- df_filled_cl %>% 
  filter(sheet %in% c("hh_roster"), !is.na(index))

# updating the main dataset with new columns

df_data_with_added_cols_roster <- cts_add_new_sm_choices_to_data(input_df_tool_data = df_loop_r_roster %>% 
                                                                   left_join(df_tool_support_data_for_loops, by = c("_submission__uuid" = "_uuid")),
                                                                 input_df_filled_cl = df_filled_cl_roster, 
                                                                 input_df_survey = df_survey,
                                                                 input_df_choices = df_choices)

# check the cleaning log
df_cl_review <- cleaningtools::review_cleaning_log(
  raw_dataset = df_data_with_added_cols_roster,
  raw_data_uuid_column = "_submission__uuid",
  cleaning_log = df_filled_cl_roster,
  cleaning_log_change_type_column = "change_type",
  change_response_value = "change_response",
  cleaning_log_question_column = "question",
  cleaning_log_uuid_column = "uuid",
  cleaning_log_new_value_column = "new_value")

# filter log for cleaning
df_final_cleaning_log_roster <- df_filled_cl_roster %>% 
  filter(!question %in% c("duration_audit_sum_all_ms", "duration_audit_sum_all_minutes", "phone_consent",
                          "_index", "_parent_index"), 
         !uuid %in% c("all")) %>% 
  filter(!str_detect(string = question, pattern = "\\w+\\/$"))

# create the clean data from the raw data and cleaning log
df_cleaning_step_roster <- cleaningtools::create_clean_data(
  raw_dataset = df_data_with_added_cols_roster %>% 
    mutate(cleaning_uuid = paste0(`_submission__uuid`, `_index`)),
  raw_data_uuid_column = "cleaning_uuid",
  cleaning_log = df_final_cleaning_log_roster %>% 
    mutate(log_cleaning_uuid = paste0(uuid, index)),
  cleaning_log_change_type_column = "change_type",
  change_response_value = "change_response",
  NA_response_value = "blank_response",
  no_change_value = "no_action",
  remove_survey_value = "remove_survey",
  cleaning_log_question_column = "question",
  cleaning_log_uuid_column = "log_cleaning_uuid",
  cleaning_log_new_value_column = "new_value")

# handle parent question columns
df_updating_sm_parents_roster <- cts_update_sm_parent_cols(input_df_cleaning_step_data = df_cleaning_step_roster,
                                                           input_uuid_col = "_submission__uuid",
                                                           input_enumerator_id_col = "enumerator_id",
                                                           input_point_id_col = "point_number",
                                                           input_collected_date_col = "today",
                                                           input_location_col = "location", 
                                                           input_dataset_type = "loop", 
                                                           input_sheet_name = "hh_roster", 
                                                           input_index_col = "_index")



# export datasets ---------------------------------------------------------

list_of_datasets <- list("raw_data" = df_tool_data %>% select(-any_of(cols_to_remove)),
                         "raw_roster" = df_loop_r_roster,
                         "cleaned_data" = df_updating_sm_parents$updated_sm_parents %>% 
                           filter(!`_uuid` %in% df_remove_survey_cl$uuid),
                         "cleaned_roster" = df_updating_sm_parents_roster$updated_sm_parents %>% 
                           filter(!`_submission__uuid` %in% df_remove_survey_cl$uuid) %>% 
                           select(-cleaning_uuid)
)

openxlsx::write.xlsx(list_of_datasets,
                     paste0("outputs/", butteR::date_file_prefix(), "_diagnostic_of_informality_data.xlsx"),
                     overwrite = TRUE)

# export datasets ---------------------------------------------------------

list_of_extra_logs <- list("extra_log_sm_parents" = df_updating_sm_parents$extra_log_sm_parents,
                           "extra_log_sm_parents_roster" = df_updating_sm_parents_roster$extra_log_sm_parents
)

openxlsx::write.xlsx(list_of_extra_logs,
                     paste0("outputs/", butteR::date_file_prefix(), "_extra_sm_parent_changes_checks_diagnostic_of_informality.xlsx"),
                     overwrite = TRUE)
