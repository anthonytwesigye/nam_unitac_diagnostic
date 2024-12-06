library(tidyverse)
library(openxlsx)

# global options can be set to further simplify things
options("openxlsx.borderStyle" = "thin")
options("openxlsx.withFilter" = FALSE)

# tool
loc_tool <- "inputs/Diagnostic_of_informality_Tool.xlsx"
df_survey <- readxl::read_excel(loc_tool, sheet = "survey")
df_choices <- readxl::read_excel(loc_tool, sheet = "choices") %>% 
  select(list_name, choice_name = name,   choice_label =label)

# extract select types
df_tool_select_type <- df_survey %>% 
  select(type, qn_name = name, label) %>% 
  filter(str_detect(string = type, pattern = "integer|date|select_one|select_multiple")) %>% 
  separate(col = type, into = c("select_type", "list_name"), sep =" ", remove = TRUE, extra = "drop" )

# tool support
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
  "undertaking_a_task",                                    "Undertaking a task",
  "i.type_of_work_done",                                    "DP1.7 Type of main work done - Sectors"
)

df_tool_support <- df_survey %>% 
  select(type, qn_name = name, label) %>% 
  filter(str_detect(string = type, pattern = "integer|date|select_one|select_multiple")) %>% 
  select(qn_name, label) %>% 
  bind_rows(df_added_composite_qns)

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
  "i.type_of_work_done",       "98 -  Other"
)

df_added_cats_extract <- df_added_categories %>% 
  mutate(name = str_extract(string = new_category, pattern = "^[0-9]{1,2}"),
         choice_label = str_replace(string = new_category, pattern = "^[0-9]{1,2}\\s\\-\\s", ""),
         survey_choice_id = paste0(parent_qn, "_", name)) %>% 
  select(survey_choice_id, choice_label)

# extract choice ids and labels
df_choices_support <- df_choices %>% 
  left_join(df_tool_select_type) %>% 
  unite("survey_choice_id", qn_name, choice_name, sep = "_", remove = FALSE) %>% 
  select(survey_choice_id, choice_label) %>% 
  bind_rows(df_added_cats_extract)




# analysis ----------------------------------------------------------------
# read analysis

df_unformatted_analysis <- read_csv("outputs/non_formatted_analysis_nam_diagnostic.csv") %>% 
    mutate(# group_var = str_replace_all(string = group_var, pattern = "%/%",replacement = "x"),
    # group_var_value = str_replace_all(string = group_var_value, pattern = "%/%",replacement = "x"),
    # int.analysis_var = ifelse(str_detect(string = analysis_var, pattern = "^i\\."), str_replace(string = analysis_var, pattern = "^i\\.", replacement = ""), analysis_var),
    int.analysis_var = analysis_var,
    analysis_choice_id = paste0(int.analysis_var, "_", analysis_var_value),
    analysis_var_value_label = ifelse(analysis_choice_id %in% df_choices_support$survey_choice_id, recode(analysis_choice_id, !!!setNames(df_choices_support$choice_label, df_choices_support$survey_choice_id)), analysis_var_value),
    Indicator = ifelse(int.analysis_var %in% df_tool_support$qn_name, recode(int.analysis_var, !!!setNames(df_tool_support$label, df_tool_support$qn_name)), int.analysis_var),
  )

# create wide analysis
df_analysis_wide <- df_unformatted_analysis %>% 
  select(-c(group_var, stat_low, stat_upp, 
            n_total, n_w, n_w_total, analysis_key)) %>% 
  # n_total, n_w, n_w_total, analysis_key)) %>% 
  mutate(group_var_value = ifelse(is.na(group_var_value), "total", group_var_value)) %>% 
  pivot_wider(names_from = c(group_var_value), values_from = c(stat, n)) %>% 
  # pivot_wider(names_from = c(group_var_value, population), values_from = c(stat, n)) %>% 
  mutate(row_id = row_number()) %>% 
  relocate(dataset, .after = "n_total")

# create order for columns (stat and n)
df_cols_for_ordering <- tibble("result_col" = df_analysis_wide %>% select(starts_with("stat_")) %>% colnames(),
                               "n_unweighted" = df_analysis_wide %>% select(starts_with("n_")) %>% colnames())

reordered_columns <- df_cols_for_ordering %>%
  pivot_longer(cols = c(result_col, n_unweighted), names_to = "entries", values_to = "columns") %>%
  pull(columns)

# reorder
extra_cols_for_analysis_tables <- c("Indicator")

df_analysis_wide_reodered <- df_analysis_wide %>%
  # relocate(any_of(extra_cols_for_analysis_tables), .before = "analysis_var") %>% 
  relocate("analysis_var_value_label", .after = "analysis_var_value") %>% 
  relocate(any_of(reordered_columns), .after = "analysis_var_value_label") %>% 
  relocate(analysis_type, .after = "analysis_var_value_label")


cols_for_num_pct_formatting <- df_analysis_wide_reodered %>% 
  select(stat_total:row_id, - any_of(c("Indicator", "dataset"))) %>% 
  select(!matches("^n_"), -row_id) %>% 
  colnames()

# extract header data

df_to_extract_header = df_analysis_wide_reodered %>% 
  select(-any_of(c("analysis_var", "analysis_var_value", "analysis_var_value_label",
                   "int.analysis_var", "analysis_choice_id", "Indicator", "row_id"))) %>% 
  colnames()

df_extracted_header_data <- tibble("old_cols" = df_to_extract_header) %>% 
  mutate("new_cols" = paste0("x", row_number())) %>% 
  mutate(old_cols = str_replace(string = old_cols, pattern = "stat_", replacement = "")) %>%
  mutate(old_cols = str_replace(string = old_cols, pattern = "^n_.+", replacement = "n")) %>% 
  mutate(old_cols = str_replace(string = old_cols, pattern = "%/%", replacement = "")) %>% 
  pivot_wider(names_from = new_cols, values_from = old_cols)

df_extracted_header <- bind_rows(df_extracted_header_data) %>% 
  mutate(x1 = "Analysis Type")

# create workbook ---------------------------------------------------------

wb <- createWorkbook()

hs1 <- createStyle(fgFill = "#049dd9", halign = "CENTER", textDecoration = "", fontColour = "white", fontSize = 12, wrapText = T, 
                   border = "TopBottomLeftRight", borderStyle = "thin", borderColour = "#000000")
hs2 <- createStyle(fgFill = "grey", halign = "LEFT", textDecoration = "Bold", fontColour = "white", wrapText = F)
hs2_no_bold <- createStyle(fgFill = "grey", halign = "LEFT", textDecoration = "", fontColour = "white", wrapText = F)
hs2_relevant <- createStyle(fgFill = "grey", halign = "LEFT", textDecoration = "", fontColour = "#808080", wrapText = F)
hs3 <- createStyle(fgFill = "#EE5859", halign = "CENTER", fontColour = "white", textDecoration = "Bold", 
                   border = "TopBottomLeftRight", borderStyle = "medium", borderColour = "#000000")

modifyBaseFont(wb = wb, fontSize = 12, fontName = "Arial Narrow")

# numbers
number_2digit_style <- openxlsx::createStyle(numFmt = "0.00")
number_1digit_style <- openxlsx::createStyle(numFmt = "0.0")
number_style <- openxlsx::createStyle(numFmt = "0")
# percent
pct = createStyle(numFmt="0.0%") # not working

addWorksheet(wb, sheetName="Diagnostic")

# header showing results headings
writeData(wb, sheet = "Diagnostic", df_extracted_header %>% head(1), startCol = 2, 
          startRow = 1, headerStyle = hs2, colNames = FALSE, 
          borders = "all", borderColour = "#000000", borderStyle = "thin")
addStyle(wb, sheet = "Diagnostic", hs1, rows = 1, cols = 1:65, gridExpand = TRUE)

setColWidths(wb = wb, sheet = "Diagnostic", cols = 1, widths = 70)
setColWidths(wb = wb, sheet = "Diagnostic", cols = 2, widths = 26)
setColWidths(wb = wb, sheet = "Diagnostic", cols = 3:65, widths = 10)

# split variables to be written in different tables with in a sheet
sheet_variables_data <- split(df_analysis_wide_reodered, factor(df_analysis_wide_reodered$analysis_var, levels = unique(df_analysis_wide_reodered$analysis_var)))

previous_row_end <- 1

for (i in 1:length(sheet_variables_data)) {
  
  current_variable_data <- sheet_variables_data[[i]]
  
  get_question <- current_variable_data %>% select(analysis_var) %>% unique() %>% pull()
  get_question_label <- current_variable_data %>% select(Indicator) %>% unique() %>% pull()
  get_qn_type <- current_variable_data %>% select(analysis_type) %>% unique() %>% pull()
  get_dataset_type <- current_variable_data %>% select(dataset) %>% unique() %>% pull()
  
  print(paste0("analysis type: ", get_qn_type))
  
  if(get_qn_type %in% c("prop_select_one", "prop_select_multiple")){
    for(n in cols_for_num_pct_formatting){class(current_variable_data[[n]])= "percentage"}
  }else{
    for(n in cols_for_num_pct_formatting){class(current_variable_data[[n]])= "numeric"}
  }
  
  # this controls rows between questions
  current_row_start <- previous_row_end + 2
  
  # print(paste0("current start row: ", current_row_start, ", variable: ", get_question))
  
  # add header for variable
  writeData(wb, sheet = "Diagnostic", get_question_label, startCol = 1, startRow = previous_row_end + 1)
  writeData(wb, sheet = "Diagnostic", get_qn_type, startCol = 2, startRow = previous_row_end + 1)
  writeData(wb, sheet = "Diagnostic", get_dataset_type, startCol = 65, startRow = previous_row_end + 1)
  addStyle(wb, sheet = "Diagnostic", hs2, rows = previous_row_end + 1, cols = 1:65, gridExpand = TRUE)
  
  # current_data_length <- max(current_variable_data$row_id) - min(current_variable_data$row_id)
  current_data_length <- nrow(current_variable_data)
  
  # print(paste0("start row: ", current_row_start, ", previous end: ", previous_row_end, ", data length: ", current_data_length, ", variable: ", get_question))
  
  writeData(wb = wb, 
            sheet = "Diagnostic", 
            x = current_variable_data %>% 
              select(-any_of(c("analysis_var", "int.analysis_var", "analysis_var_value",
                        "analysis_choice_id", "Indicator",
                        "row_id", "dataset"))
              ) %>% 
              mutate(analysis_type = NA_character_) %>% 
              arrange(desc(stat_total)), 
            startRow = current_row_start, 
            startCol = 1, 
            colNames = FALSE)
  
  previous_row_end <- current_row_start + current_data_length
}

# freeze pane
freezePane(wb, "Diagnostic", firstActiveRow = 2, firstActiveCol = 3)

# openXL(wb)

saveWorkbook(wb, paste0("outputs/", butteR::date_file_prefix(),"_formatted_analysis_nam_diagnostic.xlsx"), overwrite = TRUE)
openXL(file = paste0("outputs/", butteR::date_file_prefix(),"_formatted_analysis_nam_diagnostic.xlsx"))
