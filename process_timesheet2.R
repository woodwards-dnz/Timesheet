
# timesheet analysis ####

suppressMessages({suppressWarnings({
  library(readxl)
  library(janitor)
  library(dplyr)
  library(tidyr)
  library(lubridate)
  library(ggplot2)
  library(ggthemes)
})})


# options ####
path <- "Timesheet2022.xlsx"
done <- ymd("2022-05-16") # monday
wdays <- wday(done + days(0:6), week_start = 1, label = TRUE, abbr = TRUE)


# read timesheet ####
charges <- read_excel(path, sheet = "Charges", skip = 1) %>% clean_names() %>% drop_na(date)


# weekly and running totals ####
weekly <- charges %>% 
  arrange(period) %>% 
  group_by(period) %>% 
  summarise(
    weekhours = sum(hours),
    weektarget = median(length),
  ) %>% 
  ungroup() %>% 
  mutate(
    balance = cumsum(weekhours) - (seq(n()) - 1) * weektarget * 5  
  ) 

ggplot(weekly) +
  labs(fill = "", colour = "", x = "Week", y = "Hours") +
  geom_col(aes(x = period, y = weekhours, fill = "Week Hours")) +
  geom_hline(yintercept = seq(7.2,36,7.2), linetype = 2) +
  geom_line(aes(x = period, y = balance, colour = "Balance"), size = 1.5) +
  geom_point(aes(x = period, y = balance, colour = "Balance", fill = "Balance"), size = 4) +
  scale_x_datetime(date_breaks = "1 weeks", date_labels = "%e %b") +
  guides(colour = "none") +
  theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
  scale_y_continuous(breaks = seq(7.2,36,7.2))

print(last_plot())


# read projects ####
projects <- read_excel(path, sheet = "Projects", skip = 1) %>% clean_names()


# timesheet ####
timesheet <- left_join(weekly, projects, by = c("project", "subproject")) %>% 
  mutate(
    date = wday(date, week_start = 1, label = TRUE, abbr = TRUE)
  ) %>% # use day of week
  group_by(period, date, project, subproject, code, subcode) %>% 
  summarise(
    hours = sum(hours),
    .groups = "keep"
  ) %>% 
  pivot_wider(names_from = "date", values_from = "hours") %>% 
  mutate(
    Total = rowSums(across(any_of(wdays)), na.rm = TRUE) 
  ) %>% 
  group_by(period) %>% 
  mutate(
    period = as_date(period),
    Week = sum(Total, na.rm = TRUE)
  ) %>% 
  ungroup() %>% 
  dplyr::filter(period > done) %>% 
  dplyr::select(-subproject, -subcode) %>% 
  arrange(period, project) 

print(timesheet)


# monthly project totals ####
monthly <- projects %>% 
  dplyr::select(project, subproject, startdate, enddate, code, subcode, budgeted, profile) %>% 
  mutate(
    code2 = paste(code, pmax(subcode, "", na.rm = TRUE)),
    project2 = paste(project, subproject),
    budgeted = if_else(is.na(budgeted), 0, budgeted),
  ) %>% 
  dplyr::select(project2, startdate, enddate, code2, budgeted, profile)
  
charges2 <- charges %>%
  dplyr::select(date, project, subproject, hours) %>% 
  mutate(
    project2 = paste(project, subproject),
    month = dmy(paste("1", month(date), year(date)))
  ) %>% 
  group_by(month, project2) %>% 
  summarise(
    hours = sum(hours, na.rm = TRUE),
  ) %>% 
  ungroup() %>% 
  complete(month, project2) 

monthly2 <- left_join(monthly, charges2) %>% 
  complete(project2, month) %>% 
  dplyr::filter(!is.na(month)) %>% # drop unused codes
  dplyr::select(-startdate, -enddate) %>% 
  group_by(month) %>% 
  mutate(
    profile = median(profile, na.rm = TRUE)
  ) %>% 
  group_by(month, project2, code2) %>% 
  summarise(
    budgeted = pmax(0, median(budgeted, na.rm = TRUE), na.rm = TRUE),
    hours = pmax(0, sum(hours, na.rm = TRUE)),
    profile = median(profile, na.rm = TRUE),
  ) %>% 
  arrange(project2, code2, month) %>% 
  mutate(
    code2 = if_else(is.na(code2), "No", code2),
  ) %>% 
  
    

  
# monthly project totals 
# FIXME not correct ####
temp2 <- combined %>% 
  mutate(
    date = if_else(month(date) == 2 & year(date) == 2022, dmy_hm("1 03 2022 0:00"), date), # adjust Feb 22
    month = dmy(paste(1, month(date), year(date))),
    target = 101 * length / 7.2, # FIXME this formula is probably wrong ####
  ) %>% 
  complete(project, month, fill = list(hours = 0)) %>% 
  drop_na(month) %>% 
  group_by(project, month) %>% 
  summarise(
    monthbudget = median(budgeted, na.rm = TRUE),
    monthhours = sum(hours, na.rm = TRUE),
    monthtarget = median(target, na.rm = TRUE),
  ) %>% 
  group_by(project) %>% 
  mutate(
    monthbudget = pmax(0,median(monthbudget, na.rm = TRUE), na.rm = TRUE),
    monthtarget = pmax(0,median(monthtarget, na.rm = TRUE), na.rm = TRUE),
    keep = !all(monthbudget == 0 & monthhours == 0),
  ) %>% 
  ungroup() %>% 
  arrange(month, desc(monthbudget)) %>% 
  mutate(project = factor(project, levels = unique(project))) %>% 
  group_by(month) %>% 
  mutate(
    nonproj = project %in% c("Admin", "Leave"),
    scale = monthtarget / sum(monthbudget[!nonproj]),
    monthbudget2 = ifelse(nonproj, monthbudget, monthbudget * scale),
  ) %>% 
  ungroup()
temp2 %>% 
  dplyr::filter(keep) %>% 
  ggplot() +
  theme_grey() +
  labs(fill = "Project hours", colour = "", x = "Month", y = "Hours") +
  geom_col(aes(x = month, y = monthhours, fill = project), position = "dodge") +
  geom_col(aes(x = month, y = monthbudget2, fill = project), alpha = 0, colour = "black", position = "dodge") +
  scale_x_date(date_breaks = "1 months", date_labels = "%e %b") +
  guides(colour = "none") +
  theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1))

print(last_plot())

