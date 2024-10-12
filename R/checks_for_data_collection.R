library(tidyverse)
library(cleaningtools)
library(httr)
library(sf)
library(glue)
library(supporteR)
library(openxlsx)

source("R/support_functions.R")
source("support_files/credentials.R")


# read data ---------------------------------------------------------------

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

# joining roster loop to main shet
df_repeat_hh_roster_data <- df_tool_data %>% 
  left_join(df_loop_r_roster, by = c("_uuid" = "_submission__uuid"))


# download audit files
download_audit_files(df = df_tool_data,
                     uuid_column = "_uuid",
                     audit_dir = "inputs/audit_files",
                     usr = user_acc,
                     pass = user_pss)
# zip audit files folder
if (dir.exists("inputs/audit_files")) {
  zip::zip(zipfile = "inputs/audit_files.zip",
           files = list.dirs(path = "inputs/audit_files/", full.names = TRUE, recursive = FALSE),
           mode = "cherry-pick")
}


# cleaningtools checks ----------------------------------------------------


# check pii

pii_cols <- c("telephone","contact","name","gps","latitude","logitude","contact","geopoint")

pii_from_data <- cleaningtools::check_pii(dataset = df_tool_data, element_name = "checked_dataset", uuid_column = "_uuid")
pii_from_data$potential_PII

# duration
# read audit file
audit_list_data <- cleaningtools::create_audit_list(audit_zip_path = "inputs/audit_files.zip")
# add duration from audit
df_tool_data_with_audit_time <- cleaningtools::add_duration_from_audit(df_tool_data, uuid_column = "_uuid", audit_list = audit_list_data)

# outliers columns not to check
outlier_cols_not_4_checking <- df_tool_data %>% 
  select(matches("geopoint|gps|_index|_submit|submission|_sample_|^_id$")) %>% 
  colnames()


# check value as the initial check
list_log <- df_tool_data_with_audit_time %>%
  check_value(uuid_column = "_uuid", values_to_look = c(99, 999, 9999))

# check duration
df_check_duration <- check_duration(dataset = df_tool_data_with_audit_time %>% 
                                      filter(location_type %in% c("interview_site")),
                                    column_to_check = "duration_audit_sum_all_minutes",
                                    uuid_column = "_uuid",
                                    log_name = "duration_log",
                                    lower_bound = 10,
                                    higher_bound = 120)
  
list_log$duration_log <- df_check_duration$duration_log
  

# other logical checks ----------------------------------------------------

# outliers
numeric_cols_main_data = c("count_hh_number", "next_resp_age", "how_long_operating_enterprise", 
                           "num_people_with_similar_job", "gross_salary", "days_worked_in_week", 
                           "hours_worked_per_day", "num_employed_paid_family", "avg_salary_employed_paid_family", 
                           "num_employed_unpaid_family", "num_employed_paid_non_family", 
                           "avg_salary_employed_paid_non_family", "estimated_monthly_sales", 
                           "estimated_monthly_profit", "enterprise_age", "monthly_value_of_sales_of_enterprise", 
                           "monthly_value_of_profit_of_enterprise", "number_of_work_days", 
                           "number_of_work_hours_per_day", "monthly_value_of_sales_of_business", 
                           "monthly_value_of_profit_of_business", "number_of_work_days_a_week_contributing_worker", 
                           "work_hours_per_day_contributing_worker")

df_check_outliers <- check_outliers(dataset = df_tool_data_with_audit_time, 
                                    uuid_column = "_uuid", sm_separator = "/",
                                    strongness_factor = 3, columns_not_to_check = outlier_cols_not_4_checking)

list_log$outliers_main_data <- df_check_outliers$potential_outliers %>% 
  filter(question %in% numeric_cols_main_data)

# soft duplicates
df_check_soft_duplicates <-  check_soft_duplicates(dataset = df_tool_data_with_audit_time, kobo_survey = df_survey,
                                                   uuid_column = "_uuid",
                                                   idnk_value = "dk",
                                                   sm_separator = "/",
                                                   log_name = "soft_duplicate_log",
                                                   threshold = 25,
                                                   return_all_results = FALSE)

# missing location
# opuwo, town_council
missing_opuwo <- c("9dcf420a-19ab-48fb-b704-895da0861c31",
                   "bcfcec7b-23bf-4710-95b3-2be53b7b1613",
                   "3ab2d613-46a4-4df5-96a1-9a1dcb894661",
                   "48e05e94-df31-4e2e-a36d-d741111b8a88") 

# okangwati, settlement
missing_okangwati <- c("bd256c98-0794-493b-8c72-dba0a59e0342",
                       "80b5ccfe-7cd5-4c23-95a5-49f73d7a7c37",
                       "a1757d14-ba10-464a-a1b0-8189059c1bff",
                       "16afa6f1-a01e-440f-880f-0b76e2c73eb4")

df_missing_loc_level <- df_tool_data %>% 
  filter(is.na(interview_loc_level), !consent %in% c("2")) %>% 
  mutate(i.check.uuid =  `_uuid`,
         i.check.change_type = "change_response",
         i.check.question = "",     
         i.check.old_value =  "", 
         i.check.new_value = "",     
         i.check.issue = "missing_location",       
         i.check.other_text = "",
         i.check.comment = "",
         i.check.reviewed = "1",
         i.check.so_sm_choices = "") %>% 
  slice(rep(1:n(), each = 2)) %>%  
  group_by(i.check.uuid, i.check.change_type,  i.check.question,  i.check.old_value) %>%  
  mutate(rank = row_number(),
         i.check.question = case_when(rank == 1 ~ "interview_loc_level", 
                                      rank == 2 & i.check.uuid %in% c(missing_opuwo) ~ "town_council",
                                      rank == 2 & i.check.uuid %in% c(missing_okangwati) ~ "settlement",
         ),
         i.check.new_value = case_when(rank == 1 & i.check.uuid %in% c(missing_opuwo) ~ "town_council",
                                       rank == 1 & i.check.uuid %in% c(missing_okangwati) ~ "settlement",
                                       rank == 2 & i.check.uuid %in% c(missing_opuwo) ~ "opuwo",
                                       rank == 2 & i.check.uuid %in% c(missing_okangwati) ~ "okangwati",
         )
  ) %>% 
  batch_select_rename()
list_log$missing_loc <- df_missing_loc_level

# testing data
df_testing_data <- df_tool_data %>% 
  filter(today < as_date("2024-10-07")) %>% 
  mutate(i.check.uuid =  `_uuid`,
         i.check.change_type = "remove_survey",
         i.check.question = "",     
         i.check.old_value =  as.character(today), 
         i.check.new_value = "",     
         i.check.issue = "testing_data",       
         i.check.other_text = "",
         i.check.comment = "",
         i.check.reviewed = "1",
         i.check.so_sm_choices = "") %>% 
  batch_select_rename()
list_log$testing_data <- df_testing_data

# no consent
df_no_consent <- df_tool_data %>% 
  filter(consent == 2) %>% 
  mutate(i.check.uuid =  `_uuid`,
         i.check.change_type = "remove_survey",
         i.check.question = "",     
         i.check.old_value =  as.character(consent), 
         i.check.new_value = "",     
         i.check.issue = "no_consent",       
         i.check.other_text = "",
         i.check.comment = "",
         i.check.reviewed = "1",
         i.check.so_sm_choices = "") %>% 
  batch_select_rename()
list_log$no_consent <- df_no_consent

# incomplete_surveys
df_incomplete_surveys <- df_tool_data %>% 
  filter(location_type %in% c("interview_site"), !interview_status %in% c("1"), consent %in% c("1")) %>% 
  mutate(i.check.uuid =  `_uuid`,
         i.check.change_type = "remove_survey",
         i.check.question = "",     
         i.check.old_value =  as.character(interview_status), 
         i.check.new_value = "",     
         i.check.issue = "incomplete_surveys",       
         i.check.other_text = "",
         i.check.comment = "",
         i.check.reviewed = "1",
         i.check.so_sm_choices = "") %>% 
  batch_select_rename()
list_log$incomplete_surveys <- df_incomplete_surveys

# poor_gps_accuracy
df_poor_gps_accuracy <- df_tool_data %>% 
  filter(`_geopoint_precision` > 20) %>% 
  mutate(i.check.uuid =  `_uuid`,
         i.check.change_type = "remove_survey",
         i.check.question = "",     
         i.check.old_value =  as.character(`_geopoint_precision`), 
         i.check.new_value = "",     
         i.check.issue = "poor_gps_accuracy",       
         i.check.other_text = "",
         i.check.comment = "",
         i.check.reviewed = "",
         i.check.so_sm_choices = "") %>% 
  batch_select_rename()
list_log$poor_gps_accuracy <- df_poor_gps_accuracy

# re routing: do_you_receive
df_receive_regular_salary <- df_tool_data %>% 
  filter(do_you_receive %in% c("1")) %>% 
  mutate(i.check.uuid =  `_uuid`,
         i.check.change_type = "remove_survey",
         i.check.question = "",     
         i.check.old_value =  as.character(do_you_receive), 
         i.check.new_value = "",     
         i.check.issue = "receive_regular_salary",       
         i.check.other_text = "",
         i.check.comment = "",
         i.check.reviewed = "",
         i.check.so_sm_choices = "") %>% 
  batch_select_rename()
list_log$receive_regular_salary <- df_receive_regular_salary

# re routing: responsibility_for_business
df_responsibility_for_business <- df_tool_data %>% 
  filter(responsibility_for_business %in% c("1")) %>% 
  mutate(i.check.uuid =  `_uuid`,
         i.check.change_type = "remove_survey",
         i.check.question = "",     
         i.check.old_value =  as.character(responsibility_for_business), 
         i.check.new_value = "",     
         i.check.issue = "responsibility_for_business",       
         i.check.other_text = "",
         i.check.comment = "",
         i.check.reviewed = "",
         i.check.so_sm_choices = "") %>% 
  batch_select_rename()
list_log$responsibility_for_business <- df_responsibility_for_business


# other checks ------------------------------------------------------------

df_other_checks <- scto_other_specify(input_tool_data = df_tool_data, 
                                      input_uuid_col = "_uuid", 
                                      input_survey = df_survey, 
                                      input_choices = df_choices)
list_log$other_log <- df_other_checks

# other checks roster
df_other_checks_roster <- ctso_other_specify_repeats(input_repeat_data = df_loop_r_roster, 
                                                     input_uuid_col = "_submission__uuid", 
                                                     input_survey = df_survey, 
                                                     input_choices = df_choices,
                                                     input_sheet_name = "hh_roster",
                                                     input_index_col = "_index")
list_log$other_log_roster <- df_other_checks_roster


# check duplicate uuids ---------------------------------------------------

df_duplicate_uuids <- cts_checks_duplicate_uuids(input_tool_data = df_tool_data)
list_log$duplicate_uuid_log <- df_duplicate_uuids

# loops outliers ----------------------------------------------------------

# roster
df_loop_outliers_roster_r <- cleaningtools::check_outliers(dataset = df_loop_r_roster  %>%  mutate(loop_uuid = paste0(`_submission__uuid`, " * ", `_index`)), 
                                                           uuid_column = "loop_uuid", strongness_factor = 3,
                                                           sm_separator = "/") 

df_potential_loop_outliers_roster_r <- df_loop_outliers_roster_r$potential_outliers %>% 
  separate_wider_delim(cols = uuid, delim = " * ", names = c("i.check.uuid", "index")) %>% 
  mutate(i.check.change_type = "change_response",
         i.check.question = question,
         i.check.old_value = as.character(old_value),
         i.check.new_value = "NA",
         i.check.issue = issue,
         i.check.description = "",
         i.check.other_text = "",
         i.check.comment = "",
         i.check.reviewed = "",
         i.check.so_sm_choices = "",
         i.check.sheet = "hh_roster",
         i.check.index = index) %>% 
  batch_select_rename()%>% 
  filter(question %in% c("age"))
list_log$outliers_roster_log_r <- df_potential_loop_outliers_roster_r 



# combine the checks ------------------------------------------------------

df_combined_log <- create_combined_log_keep_change_type(dataset_name = "checked_dataset", list_of_log = list_log)


# create workbook ---------------------------------------------------------
# prep data
cols_to_add_to_log <- c("enumerator_id", "today", "region", "interview_loc_level", "municipality", "town_council", "village_council", "settlement")
cols_to_drop <- c("municipality", "town_council", "village_council", "settlement")
cols_to_maintain <- c("enumerator_id", "today", "region", "interview_loc_level", "location")



tool_support <- df_combined_log$checked_dataset %>% 
  select(uuid = `_uuid`, any_of(cols_to_add_to_log)) %>% 
  mutate(location = case_when(interview_loc_level %in% c("municipality") ~ municipality,
                              interview_loc_level %in% c("town_council") ~ town_council,
                              interview_loc_level %in% c("village_council") ~ village_council,
                              interview_loc_level %in% c("settlement") ~ settlement)) %>% 
  select(-any_of(cols_to_drop))

# checked data
df_prep_checked_data <- df_combined_log$checked_dataset %>% 
  mutate(geopoint = NA_character_,	
         `_geopoint_latitude` = NA_character_,	
         `_geopoint_longitude` = NA_character_,	
         `_geopoint_altitude` = NA_character_
)
# repeat data
df_repeat_checked_data <- tool_support %>% 
  left_join(df_loop_r_roster, by = c("uuid" = "_submission__uuid"))

# log
df_prep_cleaning_log <- df_combined_log$cleaning_log %>%
  left_join(tool_support, by = "uuid") %>% 
  relocate(any_of(cols_to_maintain), .after = uuid) %>% 
  add_qn_label_to_cl(input_cl_name_col = "question",
                     input_tool = df_survey, 
                     input_tool_name_col = "name", 
                     input_tool_label_col = "label") %>% 
  filter(!(question %in% c("_index")&issue %in% c("Possible value to be changed to NA"))) %>% 
  mutate(reviewed = ifelse(question %in% c("duration_audit_sum_all_minutes"), "1", reviewed),
         change_type = ifelse(question %in% c("duration_audit_sum_all_minutes"), "remove_survey", change_type))

df_prep_soft_duplicates_log <- df_check_soft_duplicates$soft_duplicate_log %>%
  left_join(tool_support, by = "uuid") %>%
  relocate(any_of(cols_to_maintain), .after = uuid)

df_prep_readme <- tibble::tribble(
  ~change_type_validation,                       ~description,
  "change_response", "Change the response to new_value",
  "blank_response",       "Remove and NA the response",
  "remove_survey",                "Delete the survey",
  "no_action",               "No action to take."
)

wb_log <- createWorkbook()

hs1 <- createStyle(fgFill = "#7B7B7B", textDecoration = "Bold", fontName = "Arial Narrow", fontColour = "white", fontSize = 12, wrapText = F)

modifyBaseFont(wb = wb_log, fontSize = 11, fontName = "Arial Narrow")

## checked dataset 
addWorksheet(wb_log, sheetName="checked_dataset")
setColWidths(wb = wb_log, sheet = "checked_dataset", cols = 1:ncol(df_prep_checked_data), widths = 24.89)
writeDataTable(wb = wb_log, sheet = "checked_dataset", 
               x = df_prep_checked_data , 
               startRow = 1, startCol = 1, 
               tableStyle = "TableStyleLight9",
               headerStyle = hs1)
# freeze pane
freezePane(wb = wb_log, "checked_dataset", firstActiveRow = 2, firstActiveCol = 2)

## checked repeat 
addWorksheet(wb_log, sheetName="hh_roster")
setColWidths(wb = wb_log, sheet = "hh_roster", cols = 1:ncol(df_repeat_checked_data), widths = 24.89)
writeDataTable(wb = wb_log, sheet = "hh_roster", 
               x = df_repeat_checked_data , 
               startRow = 1, startCol = 1, 
               tableStyle = "TableStyleLight9",
               headerStyle = hs1)
# freeze pane
freezePane(wb = wb_log, "hh_roster", firstActiveRow = 2, firstActiveCol = 2)

## cleaning log
addWorksheet(wb_log, sheetName="cleaning_log")
setColWidths(wb = wb_log, sheet = "cleaning_log", cols = 1:ncol(df_prep_cleaning_log), widths = 24.89)
writeDataTable(wb = wb_log, sheet = "cleaning_log", 
               x = df_prep_cleaning_log , 
               startRow = 1, startCol = 1, 
               tableStyle = "TableStyleLight9",
               headerStyle = hs1)
# freeze pane
freezePane(wb = wb_log, "cleaning_log", firstActiveRow = 2, firstActiveCol = 2)

## similar surveys
addWorksheet(wb_log, sheetName="similar_surveys")
setColWidths(wb = wb_log, sheet = "similar_surveys", cols = 1:ncol(df_prep_soft_duplicates_log), widths = 24.89)
writeData(wb = wb_log, sheet = "similar_surveys",
          x = df_prep_soft_duplicates_log ,
          startRow = 1, startCol = 1,
          withFilter = TRUE,
          headerStyle = hs1)
# freeze pane
freezePane(wb = wb_log, "similar_surveys", firstActiveRow = 2, firstActiveCol = 2)

## readme
addWorksheet(wb_log, sheetName="readme")
setColWidths(wb = wb_log, sheet = "readme", cols = 1:ncol(df_prep_readme), widths = 24.89)
writeDataTable(wb = wb_log, sheet = "readme", 
               x = df_prep_readme , 
               startRow = 1, startCol = 1, 
               tableStyle = "TableStyleLight9",
               headerStyle = hs1)
# freeze pane
freezePane(wb = wb_log, "readme", firstActiveRow = 2, firstActiveCol = 2)

# active sheet

activeSheet(wb = wb_log) <- "cleaning_log"


# openXL(wb_log)

saveWorkbook(wb_log, paste0("outputs/", butteR::date_file_prefix(),"_combined_checks_nam_diagnostic.xlsx"), overwrite = TRUE)
openXL(file = paste0("outputs/", butteR::date_file_prefix(),"_combined_checks_nam_diagnostic.xlsx"))

saveWorkbook(wb_log, paste0("inputs/combined_checks_nam_diagnostic.xlsx"), overwrite = TRUE)
