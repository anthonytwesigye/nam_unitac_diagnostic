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
df_tool_select_type <- df_survey %>%
  select(type, qn_name = name, label) %>%
  filter(str_detect(string = type, pattern = "integer|decimal|date|select_one|select_multiple")) %>%
  separate(col = type, into = c("select_type", "list_name"), sep =" ", remove = TRUE, extra = "drop" )

# extract choice ids and labels
df_choices_support <- df_choices %>%
  left_join(df_tool_select_type) %>%
  unite("survey_choice_id", qn_name, choice_name, sep = "_", remove = FALSE) %>%
  select(list_name, survey_choice_id, choice_name, choice_label)

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

df_added_composite_qns <- tibble::tribble(
  ~qn_name,                ~label,
  "i.respondent_age",   "Respondent age",
  "i.respondent_gender", "Respondent gender",
  "watching",                                              "Watching",
  "listening",                                             "Listening",
  "speaking",                                              "Speaking",
  "walking",                                               "Walking",
  "moving_around_using_equipment", "Moving around using equipment (e.g. wheelchair, etc.)",
  "lifting_and_carrying_objects",                          "Lifting and carrying objects",
  "calculating",                                           "Calculating",
  "undertaking_a_task",                                    "Undertaking a task"
)

# general columns
select_general_cols <- df_tool_select_type %>%
  filter(select_type %in% c("select_one", "integer", "decimal", "text", "date", "dateTime", "datetime"))

# select multiple cols
df_sm_choices_support <- df_choices %>%
  left_join(df_tool_select_type) %>%
  unite("survey_choice_id", qn_name, choice_name, sep = "_", remove = FALSE) %>%
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


