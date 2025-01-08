library(tidyverse)
library(supporteR)
library(openxlsx)

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
df_choices <- readxl::read_excel(loc_tool, sheet = "choices") %>% 
  select(list_name, choice_name = name,   choice_label = label) %>%
  filter(!is.na(list_name))

# extract select types

df_tool_select_type_init <- df_survey %>%
  select(type, qn_name = name, label) %>%
  filter(str_detect(string = type, pattern = "integer|decimal|text|date|select_one|select_multiple")) %>%
  separate(col = type, into = c("select_type", "list_name"), sep =" ", remove = TRUE, extra = "drop" )

df_tool_select_type_added <- tibble::tribble(
  ~select_type, ~list_name, ~qn_name,                ~label,
  "integer", NA_character_, "i.respondent_age",   "Respondent age",
  "select_one", "gender_list", "i.respondent_gender", "Respondent gender",
  "select_one", "activity_limits_list", "watching",   "Watching",
  "select_one", "activity_limits_list", "listening",  "Listening",
  "select_one", "activity_limits_list", "speaking",   "Speaking",
  "select_one", "activity_limits_list", "walking",    "Walking",
  "select_one", "activity_limits_list", "moving_around_using_equipment", "Moving around using equipment (e.g. wheelchair, etc.)",
  "select_one", "activity_limits_list", "lifting_and_carrying_objects",   "Lifting and carrying objects",
  "select_one", "activity_limits_list", "calculating",                    "Calculating",
  "select_one", "activity_limits_list", "undertaking_a_task",             "Undertaking a task",
  "select_one", "sector_list", "i.type_of_work_done",             "DP1.7 Type of main work done - Sectors"
)
df_tool_select_type <- bind_rows(df_tool_select_type_init, df_tool_select_type_added)
  
# extract choice ids and labels

# added categories
df_added_categories <- tibble::tribble(
  ~parent_qn,                                   ~new_category,
  "barriers_adopting_digital_tools",   "6 - No barrier",
  "wc_work_location",            "7 - Home of employer",
  "reasons_not_using_bank",      "8 - Employer pays cash",
  "ow_exposure_mitigation",      "4 - Reporting to authorities",
  "hold_following_namibian_ids",                       "5 - Passport",
  "hold_following_namibian_ids",                     "6 - Driving license",
  "how_far_do_customers_travel",  "4 - Within and outside the settlement",
  "highest_educ_level",    "11 - Diploma",
  "type_of_work_done",    "24 - Bartender",
  "type_of_work_done",    "25 - Tailoring",
  "type_of_work_done",    "26 - Butcher",
  "type_of_work_done",    "27 - Security Guard",
  "type_of_work_done",    "28 - Fishing",
  "type_of_work_done",    "29 - Teaching",
  "type_of_work_done",    "30 - Cashier",
  "type_of_work_done",    "31 - Mine worker",
  "type_of_work_done",    "32 - Nail technician",
  "type_of_work_done",    "33 - Forces",
  "type_of_work_done",    "34 - Construction",
  "type_of_work_done",    "35 - Carpentery and Joinery",
  "type_of_work_done",    "36 - Photographer",
  "type_of_work_done",    "37 - Printing Shop",
  "type_of_work_done",    "38 - Genal finance",
  "type_of_work_done",    "39 - Leather works",
  "type_of_work_done",    "40 - Farmer",
  "type_of_work_done",    "41 - Builder",
  "type_of_work_done",    "42 - Factory worker",
  "type_of_work_done",    "43 - Milling",
  "type_of_work_done",    "44 - Construction worker",
  "type_of_work_done",    "45 - Electronics and Electrician",
  "type_of_work_done",    "46 - Filling station",
  "type_of_work_done",    "47 - Cobbler",
  "type_of_work_done",    "48 - Horticulture",
  "type_of_work_done",    "49 - General worker",
  "type_of_work_done",    "50 - Kapana seller",
  "type_of_work_done",    "51 - Money lending",
  "type_of_work_done",    "52 - Government employee",
  "type_of_work_done",    "53 - Boutique",
  "type_of_work_done",    "54 - Crafts and Netting",
  "type_of_work_done",    "55 - General manufacture",
  "i.respondent_gender",    "1 - Male",
  "i.respondent_gender",    "2 - Female",
  "watching",                                                  "1 -  No difficulty",
  "watching",           "2 -  Mild difficulty (presents less than 25% of the time)",
  "watching", "3 -  Moderate difficulty (presents between 25% and 50% of the time)",
  "watching",   "4 -  Severe difficulty (presents between 50% and 95% of the time)",
  "watching",       "5 -  Complete difficulty (presents more than 95% of the time)",
  "listening",                                                  "1 -  No difficulty",
  "listening",           "2 -  Mild difficulty (presents less than 25% of the time)",
  "listening", "3 -  Moderate difficulty (presents between 25% and 50% of the time)",
  "listening",   "4 -  Severe difficulty (presents between 50% and 95% of the time)",
  "listening",       "5 -  Complete difficulty (presents more than 95% of the time)",
  "speaking",                                                  "1 -  No difficulty",
  "speaking",           "2 -  Mild difficulty (presents less than 25% of the time)",
  "speaking", "3 -  Moderate difficulty (presents between 25% and 50% of the time)",
  "speaking",   "4 -  Severe difficulty (presents between 50% and 95% of the time)",
  "speaking",       "5 -  Complete difficulty (presents more than 95% of the time)",
  "walking",                                                  "1 -  No difficulty",
  "walking",           "2 -  Mild difficulty (presents less than 25% of the time)",
  "walking", "3 -  Moderate difficulty (presents between 25% and 50% of the time)",
  "walking",   "4 -  Severe difficulty (presents between 50% and 95% of the time)",
  "walking",       "5 -  Complete difficulty (presents more than 95% of the time)",
  "moving_around_using_equipment",                                                  "1 -  No difficulty",
  "moving_around_using_equipment",           "2 -  Mild difficulty (presents less than 25% of the time)",
  "moving_around_using_equipment", "3 -  Moderate difficulty (presents between 25% and 50% of the time)",
  "moving_around_using_equipment",   "4 -  Severe difficulty (presents between 50% and 95% of the time)",
  "moving_around_using_equipment",       "5 -  Complete difficulty (presents more than 95% of the time)",
  "lifting_and_carrying_objects",                                                  "1 -  No difficulty",
  "lifting_and_carrying_objects",           "2 -  Mild difficulty (presents less than 25% of the time)",
  "lifting_and_carrying_objects", "3 -  Moderate difficulty (presents between 25% and 50% of the time)",
  "lifting_and_carrying_objects",   "4 -  Severe difficulty (presents between 50% and 95% of the time)",
  "lifting_and_carrying_objects",       "5 -  Complete difficulty (presents more than 95% of the time)",
  "calculating",                                                  "1 -  No difficulty",
  "calculating",           "2 -  Mild difficulty (presents less than 25% of the time)",
  "calculating", "3 -  Moderate difficulty (presents between 25% and 50% of the time)",
  "calculating",   "4 -  Severe difficulty (presents between 50% and 95% of the time)",
  "calculating",       "5 -  Complete difficulty (presents more than 95% of the time)",
  "undertaking_a_task",                                                  "1 -  No difficulty",
  "undertaking_a_task",           "2 -  Mild difficulty (presents less than 25% of the time)",
  "undertaking_a_task", "3 -  Moderate difficulty (presents between 25% and 50% of the time)",
  "undertaking_a_task",   "4 -  Severe difficulty (presents between 50% and 95% of the time)",
  "undertaking_a_task",       "5 -  Complete difficulty (presents more than 95% of the time)",
  "i.type_of_work_done",       "1 -  Food",
  "i.type_of_work_done",       "2 -  Agric, forestry & fishing",
  "i.type_of_work_done",       "3 -  Retail trade",
  "i.type_of_work_done",       "4 -  Service-oriented",
  "i.type_of_work_done",       "5 -  Transport",
  "i.type_of_work_done",       "6 -  Mining",
  "i.type_of_work_done",       "7 -  Construction",
  "i.type_of_work_done",       "8 -  Teaching",
  "i.type_of_work_done",       "9 -  Manufacturing",
  "i.type_of_work_done",       "10 -  Financial activities",
  "i.type_of_work_done",       "98 -  Other"
)

df_choices_added <- df_added_categories %>% 
  mutate(choice_name = str_extract(string = new_category, pattern = "^[0-9]{1,2}"),
         choice_label = str_replace(string = new_category, pattern = "^[0-9]{1,2}\\s\\-\\s", ""),
         qn_name = parent_qn) %>% 
  left_join(df_tool_select_type) %>% 
  mutate(list_name = ifelse(qn_name %in% c("i.type_of_work_done"), "sector_list", list_name)) %>% 
  select(list_name, choice_name, choice_label)

df_choices_support <- df_choices %>%
  bind_rows(df_choices_added) %>% 
  left_join(df_tool_select_type) %>%
  unite("survey_choice_id", qn_name, choice_name, sep = "_", remove = FALSE) %>%
  select(select_type, list_name, survey_choice_id, choice_name, choice_label, qn_name, label)


# Main data ---------------------------------------------------------------

# handle select one

valid_data_cols <- df_main_clean_data %>% 
  janitor::remove_empty(which = "cols") %>% 
  colnames()
# select one cols
select_one_cols <- df_tool_select_type %>%
  filter(select_type %in% c("select_one"), qn_name %in% valid_data_cols)

df_main_clean_data_with_so_labels <- df_main_clean_data %>%
  mutate(across(.cols = any_of(select_one_cols$qn_name), .fns = ~ ifelse(!(is.na(.x)| .x %in% c("NA", "")), paste0(cur_column(), "_", .),.x))) %>%
  mutate(across(.cols = any_of(select_one_cols$qn_name), .fns = ~recode(.x, !!!setNames(df_choices_support$choice_label, df_choices_support$survey_choice_id))))


# select multiple data update

select_multiple_cols <- df_tool_select_type %>%
  filter(select_type %in% c("select_multiple"), qn_name %in% valid_data_cols)

# loop through the SM questions and data to update the SM cols
# for a given sm question,
# - create a named vector,
# - mutate and replace using the named vector
# - put escape characters while creating the named vector [\\b]

sm_extract <- select_multiple_cols %>% 
  pull(qn_name)

df_data_for_update <- df_main_clean_data_with_so_labels

for (sm_col in sm_extract) {
  current_list_name <- select_multiple_cols %>% 
    filter(qn_name %in% sm_col) %>% 
    pull(list_name)
  current_choices <- df_choices_support %>% 
    filter(list_name %in% current_list_name) %>% 
    mutate(choice_name = paste0("\\b", choice_name,"\\b")) %>% 
    select(choice_name, choice_label)
  current_choices_vec <- current_choices %>% 
    pull(choice_label) %>% 
    set_names(current_choices$choice_name)
  # print(current_choices_vec)
  
  df_current_update <- df_data_for_update %>%
    mutate(!!sm_col := str_replace_all(string = !!sym(sm_col),  current_choices_vec))
  
  df_data_for_update <- df_current_update
  
}

df_main_clean_data_with_so_sm_labels <- df_data_for_update


# Roster data ---------------------------------------------------------------

# handle select one

valid_data_cols_roster <- df_clean_loop_r_roster %>% 
  janitor::remove_empty(which = "cols") %>% 
  colnames()
# select one cols
select_one_cols_roster <- df_tool_select_type %>%
  filter(select_type %in% c("select_one"), qn_name %in% valid_data_cols_roster)

df_roster_clean_data_with_so_labels <- df_clean_loop_r_roster %>%
  mutate(across(.cols = any_of(select_one_cols_roster$qn_name), .fns = ~ ifelse(!(is.na(.x)| .x %in% c("NA", "")), paste0(cur_column(), "_", .),.x))) %>%
  mutate(across(.cols = any_of(select_one_cols_roster$qn_name), .fns = ~recode(.x, !!!setNames(df_choices_support$choice_label, df_choices_support$survey_choice_id))))


# select multiple data update

select_multiple_cols_roster <- df_tool_select_type %>%
  filter(select_type %in% c("select_multiple"), qn_name %in% valid_data_cols_roster)

# loop through the SM questions and data to update the SM cols
# for a given sm question,
# - create a named vector,
# - mutate and replace using the named vector
# - put escape characters while creating the named vector [\\b]

sm_extract_roster <- select_multiple_cols_roster %>% 
  pull(qn_name)

df_data_for_update_roster <- df_roster_clean_data_with_so_labels

for (sm_col in sm_extract_roster) {
  current_list_name_roster <- select_multiple_cols_roster %>% 
    filter(qn_name %in% sm_col) %>% 
    pull(list_name)
  current_choices_roster <- df_choices_support %>% 
    filter(list_name %in% current_list_name_roster) %>% 
    mutate(choice_name = paste0("\\b", choice_name,"\\b")) %>% 
    select(choice_name, choice_label)
  current_choices_vec_roster <- current_choices_roster %>% 
    pull(choice_label) %>% 
    set_names(current_choices_roster$choice_name)
  # print(current_choices_vec_roster)
  
  df_current_update_roster <- df_data_for_update_roster %>%
    mutate(!!sm_col := str_replace_all(string = !!sym(sm_col),  current_choices_vec_roster))
  
  df_data_for_update_roster <- df_current_update_roster
  
}

df_roster_clean_data_with_so_sm_labels <- df_data_for_update_roster


# format the headers ------------------------------------------------------

# general columns
select_general_cols <- df_tool_select_type %>%
  filter(select_type %in% c("select_one", "integer", "decimal", "text", "date", "dateTime", "datetime"))

# select multiple cols
df_sm_choices_support <- df_choices_support %>%
  filter(select_type %in% c("select_multiple"))

# individual choices handling
df_sm_choices_support_individual <- df_sm_choices_support %>%
  mutate(qn_name = paste0(qn_name, "/", choice_name),
         label = paste0(label, "/", choice_label))

# parent question handling
df_sm_choices_support_parent <- df_sm_choices_support %>%
  group_by(list_name) %>%
  filter(row_number() == 1) %>%
  mutate(survey_choice_id = NA_character_,
         choice_name = NA_character_)

df_sm_choices_support_combined <- bind_rows(df_sm_choices_support_parent, df_sm_choices_support_individual)

# extract header data main

df_to_extract_header = df_main_clean_data_with_so_sm_labels %>%
  colnames()

df_extracted_header_data <- tibble("old_cols" = df_to_extract_header) %>%
  mutate(new_cols = paste0("x", row_number()),
         labels = ifelse(old_cols %in% select_general_cols$qn_name, recode(old_cols, !!!setNames(select_general_cols$label, select_general_cols$qn_name)), old_cols),
         labels = ifelse(old_cols %in% df_sm_choices_support_combined$qn_name, recode(old_cols, !!!setNames(df_sm_choices_support_combined$label, df_sm_choices_support_combined$qn_name)), labels),
  ) %>%
  select(-old_cols) %>%
  pivot_wider(names_from = new_cols, values_from = labels)

# extract header data roster

df_to_extract_header_roster = df_roster_clean_data_with_so_sm_labels %>%
  colnames()

df_extracted_header_data_roster <- tibble("old_cols" = df_to_extract_header_roster) %>%
  mutate(new_cols = paste0("x", row_number()),
         labels = ifelse(old_cols %in% select_general_cols$qn_name, recode(old_cols, !!!setNames(select_general_cols$label, select_general_cols$qn_name)), old_cols),
         labels = ifelse(old_cols %in% df_sm_choices_support_combined$qn_name, recode(old_cols, !!!setNames(df_sm_choices_support_combined$label, df_sm_choices_support_combined$qn_name)), labels),
  ) %>%
  select(-old_cols) %>%
  pivot_wider(names_from = new_cols, values_from = labels)


# create workbook ---------------------------------------------------------
# workbook
wb_data_with_labs <- createWorkbook()

hs1 <- createStyle(fgFill = "#7B7B7B", textDecoration = "Bold", fontName = "Arial Narrow", fontColour = "white", fontSize = 12, wrapText = F)

modifyBaseFont(wb = wb_data_with_labs, fontSize = 11, fontName = "Arial Narrow")

# clean_main_data_labels
addWorksheet(wb_data_with_labs, sheetName="clean_main_data_labels")
setColWidths(wb = wb_data_with_labs, sheet = "clean_main_data_labels", cols = 1:ncol(df_main_clean_data_with_so_sm_labels), widths = 24.89)

# header
writeData(wb = wb_data_with_labs, sheet = "clean_main_data_labels", df_extracted_header_data %>% head(1), startCol = 1,
          startRow = 1, headerStyle = hs1, colNames = FALSE)
addStyle(wb = wb_data_with_labs, sheet = "clean_main_data_labels", hs1, rows = 1, cols = 1:ncol(df_main_clean_data_with_so_sm_labels), gridExpand = TRUE)

# data
writeData(wb = wb_data_with_labs, sheet = "clean_main_data_labels",
          x = df_main_clean_data_with_so_sm_labels ,
          startRow = 2, startCol = 1,
          colNames = FALSE,
          withFilter = FALSE)

# roster_data_labels
addWorksheet(wb_data_with_labs, sheetName="roster_data_labels")
setColWidths(wb = wb_data_with_labs, sheet = "roster_data_labels", cols = 1:ncol(df_roster_clean_data_with_so_sm_labels), widths = 24.89)

# header
writeData(wb = wb_data_with_labs, sheet = "roster_data_labels", df_extracted_header_data_roster %>% head(1), startCol = 1,
          startRow = 1, headerStyle = hs1, colNames = FALSE)
addStyle(wb = wb_data_with_labs, sheet = "roster_data_labels", hs1, rows = 1, cols = 1:ncol(df_roster_clean_data_with_so_sm_labels), gridExpand = TRUE)

# data
writeData(wb = wb_data_with_labs, sheet = "roster_data_labels",
          x = df_roster_clean_data_with_so_sm_labels ,
          startRow = 2, startCol = 1,
          colNames = FALSE,
          withFilter = FALSE)

# openXL(wb_data_with_labs)
saveWorkbook(wb_data_with_labs, paste0("outputs/", butteR::date_file_prefix(),"_diagnostic_of_informality_data_with_labels.xlsx"), overwrite = TRUE)
openXL(file = paste0("outputs/", butteR::date_file_prefix(),"_diagnostic_of_informality_data_with_labels.xlsx"))



