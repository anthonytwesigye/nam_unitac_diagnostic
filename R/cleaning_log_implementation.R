library(tidyverse)
library(glue)
library(cleaningtools)
library(httr)
library(supporteR)
library(openxlsx)

loc_data <- "inputs/Diagnostic_of_informality_Data.xlsx"

# main data
data_nms <- names(readxl::read_excel(path = loc_data, n_max = 2000))
c_types <- ifelse(str_detect(string = data_nms, pattern = "_other$|\\/"), "text", "guess")
df_tool_data <- readxl::read_excel(loc_data, col_types = c_types) %>% 
  mutate(point_number = paste0("pt_", row_number()),
         location = case_when(interview_loc_level %in% c("municipality") ~ municipality,
                              interview_loc_level %in% c("town_council") ~ town_council,
                              interview_loc_level %in% c("village_council") ~ village_council,
                              interview_loc_level %in% c("settlement") ~ settlement))

# loops
# roster
data_nms_r_roster <- names(readxl::read_excel(path = loc_data, n_max = 2000, sheet = "hh_roster"))
c_types_r_roster <- ifelse(str_detect(string = data_nms_r_roster, pattern = "_other$|\\/"), "text", "guess")
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
cols_to_remove <- c("deviceid", "audit", "audit_URL", 
                    "latitude", "longitude", "instance_name", 
                    "point_number", "init_resp_age", "next_resp_age", "init_resp_gender", "next_resp_gender",
                    "init_resp_labour_capacity", "valid_provider_of_the_job")

cols_to_remove_loc <- c("geopoint", "_geopoint_latitude", "_geopoint_longitude",
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

# interview site data

df_interview_site_data <- df_data_with_added_cols %>% 
  filter(location_type %in% c("interview_site"))

# filter log for cleaning
df_final_cleaning_log <- df_filled_cl_main %>% 
  filter(!question %in% c("duration_audit_sum_all_ms", "duration_audit_sum_all_minutes", "phone_consent"), 
         !uuid %in% c("all")) %>% 
  filter(!str_detect(string = question, pattern = "\\w+\\/$")) %>% 
  filter(uuid %in% df_interview_site_data$`_uuid`)

# create the clean data from the raw data and cleaning log
df_cleaning_step <- cleaningtools::create_clean_data(
  raw_dataset = df_interview_site_data,
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


# poi data ----------------------------------------------------------------

cols_to_remove_poi <- c("deviceid", "audit", "audit_URL",
                    "point_number", "location")

df_poi_data <- df_data_with_added_cols %>% 
  filter(location_type %in% c("poi"))

# filter log for cleaning
df_poi_log <- df_filled_cl_main %>% 
  filter(!question %in% c("duration_audit_sum_all_ms", "duration_audit_sum_all_minutes", "phone_consent"), 
         !uuid %in% c("all")) %>% 
  filter(uuid %in% df_poi_data$`_uuid`)


# create the clean data from the raw data and cleaning log
df_cleaning_step_poi <- cleaningtools::create_clean_data(
  raw_dataset = df_poi_data,
  raw_data_uuid_column = "_uuid",
  cleaning_log = df_poi_log,
  cleaning_log_change_type_column = "change_type",
  change_response_value = "change_response",
  NA_response_value = "blank_response",
  no_change_value = "no_action",
  remove_survey_value = "remove_survey",
  cleaning_log_question_column = "question",
  cleaning_log_uuid_column = "uuid",
  cleaning_log_new_value_column = "new_value") %>% 
  janitor::remove_empty(which = "cols") %>% 
  select(-any_of(cols_to_remove_poi))

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

df_out_raw_data <- df_tool_data %>% select(-any_of(c(cols_to_remove, cols_to_remove_loc)))

df_updated_main_data <- df_updating_sm_parents$updated_sm_parents %>% 
  mutate(i.respondent_age = ifelse(is.na(next_resp_age), init_resp_age, next_resp_age),
         i.respondent_gender = ifelse(is.na(next_resp_gender), init_resp_gender, next_resp_gender))

df_out_clean_data <- df_updated_main_data %>% 
  select(-any_of(c(cols_to_remove, cols_to_remove_loc))) %>% 
  filter(!`_uuid` %in% df_remove_survey_cl$uuid)

df_out_clean_data_with_loc <-  df_updated_main_data %>% 
  select(-any_of(cols_to_remove)) %>% 
  filter(!`_uuid` %in% df_remove_survey_cl$uuid)



df_out_clean_roster <- df_updating_sm_parents_roster$updated_sm_parents %>% 
  filter(!`_submission__uuid` %in% df_remove_survey_cl$uuid) %>% 
  select(-any_of(c("cleaning_uuid", "point_number"))) %>% 
  mutate(i.type_of_work_done = case_when(type_of_work_done %in% c("1", "4", "5") ~ "1",
                                         type_of_work_done %in% c("28") ~ "2",
                                         type_of_work_done %in% c("7", "8", "11", "12", "13", "20") ~ "3",
                                         type_of_work_done %in% c("2", "3", "6", "9", "10", "15", "16", "17", "18", "21", "24", "25", "27") ~ "4",
                                         type_of_work_done %in% c("19", "23") ~ "5",
                                         type_of_work_done %in% c("14", "22", "98") ~ "98",
                                         ))


# list_of_datasets <- list("raw_data" = df_out_raw_data,
#                          "raw_roster" = df_loop_r_roster,
#                          "cleaned_data" = df_out_clean_data,
#                          "cleaned_roster" = df_out_clean_roster,
#                          "cleaned_poi" = df_cleaning_step_poi
# )
# 
# openxlsx::write.xlsx(list_of_datasets,
#                      paste0("outputs/", butteR::date_file_prefix(), "_diagnostic_of_informality_data.xlsx"),
#                      overwrite = TRUE)
# 
# # export datasets ---------------------------------------------------------
# 
# list_of_extra_logs <- list("extra_log_sm_parents" = df_updating_sm_parents$extra_log_sm_parents,
#                            "extra_log_sm_parents_roster" = df_updating_sm_parents_roster$extra_log_sm_parents
# )
# 
# openxlsx::write.xlsx(list_of_extra_logs,
#                      paste0("outputs/", butteR::date_file_prefix(), "_extra_sm_parent_changes_checks_diagnostic_of_informality.xlsx"),
#                      overwrite = TRUE)


# format export -----------------------------------------------------------

df_added_categories <- tibble::tribble(
  ~parent_qn,                                                                    ~question_label,                           ~new_category,
  "barriers_adopting_digital_tools", "TE1.9.2 What barriers do you face in adopting digital tools for your enterprise?",                        "6 - No barrier",
  "wc_work_location",                                                     "WC1.12 Do you mainly work in",                  "7 - Home of employer",
  "reasons_not_using_bank",                                     "WC1.7d Why are you not using a bank account?",                "8 - Employer pays cash",
  "ow_exposure_mitigation",                   "OW1.12b Have you done the following to mitigate the exposures:",          "4 - Reporting to authorities",
  "hold_following_namibian_ids",                              "OD1.1.2a You hold the following identity documents:",                          "5 - Passport",
  "hold_following_namibian_ids",                              "OD1.1.2a You hold the following identity documents:",                   "6 - Driving license",
  "how_far_do_customers_travel",       "OW1.5c How far do most of your customers travel to reach your enterprises?", "4 - Within and outside the settlement",
  "highest_educ_level",                                          "DP1.4 Highest level at school completed",                          "11 - Diploma",
  "type_of_work_done",                                                     "DP1.7 Type of main work done",                        "24 - Bartender",
  "type_of_work_done",                                                     "DP1.7 Type of main work done",                        "25 - Tailoring",
  "type_of_work_done",                                                     "DP1.7 Type of main work done",                          "26 - Butcher",
  "type_of_work_done",                                                     "DP1.7 Type of main work done",                   "27 - Security Guard",
  "type_of_work_done",                                                     "DP1.7 Type of main work done",                          "28 - Fishing",
  "i.type_of_work_done",                                                     "DP1.7 Type of main work done",                          "1 - Food",
  "i.type_of_work_done",                                                     "DP1.7 Type of main work done",                          "2 - Agric, forestry & fishing",
  "i.type_of_work_done",                                                     "DP1.7 Type of main work done",                          "3 - Retail trade",
  "i.type_of_work_done",                                                     "DP1.7 Type of main work done",                          "4 - Service-oriented",
  "i.type_of_work_done",                                                     "DP1.7 Type of main work done",                          "5 - Transport",
  "i.type_of_work_done",                                                     "DP1.7 Type of main work done",                          "98 - Other"
)

# datasets to export
# "raw_data" = df_out_raw_data,
# "raw_roster" = df_loop_r_roster,
# "cleaned_data" = df_out_clean_data,
# "cleaned_roster" = df_out_clean_roster,
# "cleaned_poi" = df_cleaning_step_poi

wb_cleaned_data <- createWorkbook()

hs1 <- createStyle(fgFill = "#7B7B7B", textDecoration = "Bold", fontName = "Arial Narrow", fontColour = "white", fontSize = 12, wrapText = F)

modifyBaseFont(wb = wb_cleaned_data, fontSize = 11, fontName = "Arial Narrow")

## raw_data
addWorksheet(wb_cleaned_data, sheetName="raw_data")
setColWidths(wb = wb_cleaned_data, sheet = "raw_data", cols = 1:ncol(df_out_raw_data), widths = 24.89)
writeDataTable(wb = wb_cleaned_data, sheet = "raw_data", 
               x = df_out_raw_data , 
               startRow = 1, startCol = 1, 
               tableStyle = "TableStyleLight9",
               headerStyle = hs1)
# freeze pane
freezePane(wb = wb_cleaned_data, "raw_data", firstActiveRow = 2, firstActiveCol = 2)

## raw_roster 
addWorksheet(wb_cleaned_data, sheetName="raw_roster")
setColWidths(wb = wb_cleaned_data, sheet = "raw_roster", cols = 1:ncol(df_loop_r_roster), widths = 24.89)
writeDataTable(wb = wb_cleaned_data, sheet = "raw_roster", 
               x = df_loop_r_roster , 
               startRow = 1, startCol = 1, 
               tableStyle = "TableStyleLight9",
               headerStyle = hs1)
# freeze pane
freezePane(wb = wb_cleaned_data, "raw_roster", firstActiveRow = 2, firstActiveCol = 2)

## cleaned_data 
addWorksheet(wb_cleaned_data, sheetName="cleaned_data")
setColWidths(wb = wb_cleaned_data, sheet = "cleaned_data", cols = 1:ncol(df_out_clean_data), widths = 24.89)
writeDataTable(wb = wb_cleaned_data, sheet = "cleaned_data", 
               x = df_out_clean_data , 
               startRow = 1, startCol = 1, 
               tableStyle = "TableStyleLight9",
               headerStyle = hs1)
# freeze pane
freezePane(wb = wb_cleaned_data, "cleaned_data", firstActiveRow = 2, firstActiveCol = 2)

## cleaned_roster 
addWorksheet(wb_cleaned_data, sheetName="cleaned_roster")
setColWidths(wb = wb_cleaned_data, sheet = "cleaned_roster", cols = 1:ncol(df_out_clean_roster), widths = 24.89)
writeDataTable(wb = wb_cleaned_data, sheet = "cleaned_roster", 
               x = df_out_clean_roster , 
               startRow = 1, startCol = 1, 
               tableStyle = "TableStyleLight9",
               headerStyle = hs1)
# freeze pane
freezePane(wb = wb_cleaned_data, "cleaned_roster", firstActiveRow = 2, firstActiveCol = 2)

## cleaned_poi 
addWorksheet(wb_cleaned_data, sheetName="cleaned_poi")
setColWidths(wb = wb_cleaned_data, sheet = "cleaned_poi", cols = 1:ncol(df_cleaning_step_poi), widths = 24.89)
writeDataTable(wb = wb_cleaned_data, sheet = "cleaned_poi", 
               x = df_cleaning_step_poi , 
               startRow = 1, startCol = 1, 
               tableStyle = "TableStyleLight9",
               headerStyle = hs1)
# freeze pane
freezePane(wb = wb_cleaned_data, "cleaned_poi", firstActiveRow = 2, firstActiveCol = 2)


## newcategories
addWorksheet(wb_cleaned_data, sheetName="newcategories")
setColWidths(wb = wb_cleaned_data, sheet = "newcategories", cols = 1:ncol(df_added_categories), widths = 24.89)
writeDataTable(wb = wb_cleaned_data, sheet = "newcategories", 
               x = df_added_categories , 
               startRow = 1, startCol = 1, 
               tableStyle = "TableStyleLight9",
               headerStyle = hs1)
# freeze pane
freezePane(wb = wb_cleaned_data, "newcategories", firstActiveRow = 2, firstActiveCol = 2)

# active sheet

activeSheet(wb = wb_cleaned_data) <- "cleaned_data"


saveWorkbook(wb_cleaned_data, paste0("outputs/", butteR::date_file_prefix(),"_diagnostic_of_informality_cleaned_data.xlsx"), overwrite = TRUE)
openXL(file = paste0("outputs/", butteR::date_file_prefix(),"_diagnostic_of_informality_cleaned_data.xlsx"))

saveWorkbook(wb_cleaned_data, paste0("support_files/diagnostic_of_informality_cleaned_data.xlsx"), overwrite = TRUE)



# data with location ------------------------------------------------------

wb_cleaned_data_with_loc <- createWorkbook()

## cleaned_data with location
addWorksheet(wb_cleaned_data_with_loc, sheetName="cleaned_data_loc")
setColWidths(wb = wb_cleaned_data_with_loc, sheet = "cleaned_data_loc", cols = 1:ncol(df_out_clean_data_with_loc), widths = 24.89)
writeDataTable(wb = wb_cleaned_data_with_loc, sheet = "cleaned_data_loc", 
               x = df_out_clean_data_with_loc , 
               startRow = 1, startCol = 1, 
               tableStyle = "TableStyleLight9",
               headerStyle = hs1)
# freeze pane
freezePane(wb = wb_cleaned_data_with_loc, "cleaned_data_loc", firstActiveRow = 2, firstActiveCol = 2)

## cleaned_roster 
addWorksheet(wb_cleaned_data_with_loc, sheetName="cleaned_roster")
setColWidths(wb = wb_cleaned_data_with_loc, sheet = "cleaned_roster", cols = 1:ncol(df_out_clean_roster), widths = 24.89)
writeDataTable(wb = wb_cleaned_data_with_loc, sheet = "cleaned_roster", 
               x = df_out_clean_roster , 
               startRow = 1, startCol = 1, 
               tableStyle = "TableStyleLight9",
               headerStyle = hs1)
# freeze pane
freezePane(wb = wb_cleaned_data_with_loc, "cleaned_roster", firstActiveRow = 2, firstActiveCol = 2)

## cleaned_poi 
addWorksheet(wb_cleaned_data_with_loc, sheetName="cleaned_poi")
setColWidths(wb = wb_cleaned_data_with_loc, sheet = "cleaned_poi", cols = 1:ncol(df_cleaning_step_poi), widths = 24.89)
writeDataTable(wb = wb_cleaned_data_with_loc, sheet = "cleaned_poi", 
               x = df_cleaning_step_poi , 
               startRow = 1, startCol = 1, 
               tableStyle = "TableStyleLight9",
               headerStyle = hs1)
# freeze pane
freezePane(wb = wb_cleaned_data_with_loc, "cleaned_poi", firstActiveRow = 2, firstActiveCol = 2)

## newcategories
addWorksheet(wb_cleaned_data_with_loc, sheetName="newcategories")
setColWidths(wb = wb_cleaned_data_with_loc, sheet = "newcategories", cols = 1:ncol(df_added_categories), widths = 24.89)
writeDataTable(wb = wb_cleaned_data_with_loc, sheet = "newcategories", 
               x = df_added_categories , 
               startRow = 1, startCol = 1, 
               tableStyle = "TableStyleLight9",
               headerStyle = hs1)
# freeze pane
freezePane(wb = wb_cleaned_data_with_loc, "newcategories", firstActiveRow = 2, firstActiveCol = 2)

# active sheet

activeSheet(wb = wb_cleaned_data_with_loc) <- "cleaned_data_loc"


saveWorkbook(wb_cleaned_data_with_loc, paste0("outputs/", butteR::date_file_prefix(),"_diagnostic_of_informality_data_loc.xlsx"), overwrite = TRUE)
openXL(file = paste0("outputs/", butteR::date_file_prefix(),"_diagnostic_of_informality_data_loc.xlsx"))
