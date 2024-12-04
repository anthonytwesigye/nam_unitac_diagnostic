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
df_choices <- readxl::read_excel(loc_tool, sheet = "choices") %>% 
  select(list_name, choice_name = name,   choice_label = label) %>%
  filter(!is.na(list_name))

# extract select types
changed_to_so_cols <- c("watching",
                       "listening",
                       "speaking",
                       "walking",
                       "moving_around_using_equipment",
                       "lifting_and_carrying_objects",
                       "calculating",
                       "undertaking_a_task")
so_composite_cols <- c("i.respondent_gender")

df_tool_select_type_init <- df_survey %>%
  select(type, qn_name = name, label) %>%
  filter(str_detect(string = type, pattern = "integer|decimal|date|select_one|select_multiple")) %>%
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
  "select_one", "activity_limits_list", "undertaking_a_task",             "Undertaking a task"
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
  "type_of_work_done",    "28 - Fishing"
)

df_choices_added <- df_added_categories %>% 
  mutate(choice_name = str_extract(string = new_category, pattern = "^[0-9]{1,2}"),
         choice_label = str_replace(string = new_category, pattern = "^[0-9]{1,2}\\s\\-\\s", ""),
         qn_name = parent_qn) %>% 
  left_join(df_tool_select_type) %>% 
  select(list_name, choice_name, choice_label)

df_choices_support <- df_choices %>%
  bind_rows(df_choices_added) %>% 
  left_join(df_tool_select_type) %>%
  unite("survey_choice_id", qn_name, choice_name, sep = "_", remove = FALSE) %>%
  select(select_type, list_name, survey_choice_id, choice_name, choice_label, qn_name, label)


# handle select one -------------------------------------------------------

valid_data_cols <- df_main_clean_data %>% 
  janitor::remove_empty(which = "cols") %>% 
  colnames()
# select one cols
select_one_cols <- df_tool_select_type %>%
  filter(select_type %in% c("select_one"), qn_name %in% valid_data_cols)

df_main_clean_data_with_so_labels <- df_main_clean_data %>%
  mutate(across(.cols = any_of(select_one_cols$qn_name), .fns = ~ ifelse(!(is.na(.x)| .x %in% c("NA", "")), paste0(cur_column(), "_", .),.x))) %>%
  mutate(across(.cols = any_of(select_one_cols$qn_name), .fns = ~recode(.x, !!!setNames(df_choices_support$choice_label, df_choices_support$survey_choice_id))))



# select multiple data update ---------------------------------------------

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

# extract header data

df_to_extract_header = df_main_clean_data_with_so_sm_labels %>%
  colnames()

df_extracted_header_data <- tibble("old_cols" = df_to_extract_header) %>%
  mutate(new_cols = paste0("x", row_number()),
         labels = ifelse(old_cols %in% select_general_cols$qn_name, recode(old_cols, !!!setNames(select_general_cols$label, select_general_cols$qn_name)), old_cols),
         labels = ifelse(old_cols %in% df_sm_choices_support_combined$qn_name, recode(old_cols, !!!setNames(df_sm_choices_support_combined$label, df_sm_choices_support_combined$qn_name)), labels),
  ) %>%
  select(-old_cols) %>%
  pivot_wider(names_from = new_cols, values_from = labels)


