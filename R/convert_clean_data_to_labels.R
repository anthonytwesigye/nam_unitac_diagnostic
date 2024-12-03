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
  filter(str_detect(string = type, pattern = "integer|date|select_one|select_multiple")) %>%
  separate(col = type, into = c("select_type", "list_name"), sep =" ", remove = TRUE, extra = "drop" )

# extract choice ids and labels
df_choices_support <- df_choices %>%
  left_join(df_tool_select_type) %>%
  unite("survey_choice_id", qn_name, choice_name, sep = "_", remove = FALSE) %>%
  select(survey_choice_id, choice_label)

# handle select one -------------------------------------------------------

valid_data_cols <- df_main_clean_data %>% 
  janitor::remove_empty(which = "cols") %>% 
  colnames()
# select one cols
select_one_cols <- df_tool_select_type %>%
  filter(select_type %in% c("select_one"), qn_name %in% valid_data_cols)

df_main_clean_data_with_so_labels <- df_main_clean_data %>%
  # mutate(across(.cols = any_of(select_one_cols$qn_name), .fns = ~paste0(cur_column(), "_", .))) %>%
  mutate(across(.cols = any_of(select_one_cols$qn_name), .fns = ~ ifelse(!(is.na(.x)| .x %in% c("NA", "")), paste0(cur_column(), "_", .),.x))) %>%
  mutate(across(.cols = any_of(select_one_cols$qn_name), .fns = ~recode(.x, !!!setNames(df_choices_support$choice_label, df_choices_support$survey_choice_id))))

