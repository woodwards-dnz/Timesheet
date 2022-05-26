
# timesheet analysis ####

library(stringr)
library(readxl)
library(janitor)
library(plyr) # avoid dplyr bug
library(dplyr)
library(tidyr)
library(lubridate)
library(ggplot2)
library(ggthemes)


# options ####
path <- "Timesheet2022.xlsx"
done <- ymd("2022-05-16") # monday
wdays <- wday(done + days(0:6), week_start = 1, label = TRUE, abbr = TRUE)


# read timesheet ####
charges <- read_excel(path, sheet = "Charges", skip = 1) %>% clean_names() %>% drop_na(date) %>% 
  mutate(
    date = as_date(date),
    period = as_date(period),
    month = dmy(paste("1", month(date), year(date))),
    year = if_else(month(date) >= 6, year(date), year(date) - 1),
  ) %>% 
  dplyr::select(year, month, period, date, project, subproject, hours, length)


# read projects
projects <- read_excel(path, sheet = "Projects", skip = 1) %>% clean_names() %>% 
  mutate(
    startdate = as_date(startdate),
    enddate = as_date(enddate),
    subcode = if_else(is.na(subcode), "", as.character(subcode)),
    year = if_else(month(startdate) >= 6, year(startdate), year(startdate) - 1),
    enddate = if_else(is.na(enddate), max(enddate, na.rm = TRUE), enddate)
  ) %>% 
  dplyr::select(year, startdate, enddate, project, subproject, code, subcode, budgeted, profile)


# combine
combined <- left_join(charges, projects, by = c("year", "project", "subproject")) %>% 
  dplyr::filter(date >= startdate & date <= enddate) 


# weekly and running totals ####
weekly <- combined %>% 
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
  scale_x_date(date_breaks = "1 weeks", date_labels = "%e %b") +
  guides(colour = "none") +
  theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
  scale_y_continuous(breaks = seq(7.2,36,7.2))

print(last_plot())


# timesheet ####
timesheet <- combined %>% 
  mutate(
    wday = wday(date, week_start = 1, label = TRUE, abbr = TRUE),
    date = NULL,
  ) %>% 
  arrange(period, wday, project, subproject, code, subcode) %>% 
  group_by(period, wday, project, subproject, code, subcode) %>% 
  summarize(
    hours = sum(hours),
    .groups = "keep"
  ) %>% 
  ungroup() %>% 
  pivot_wider(id_cols = all_of(c("period", "project", "subproject", "code", "subcode")), names_from = "wday", values_from = "hours") %>% 
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
monthly <- combined %>%
  mutate(
    project2 = paste(project, subproject),
    code2 = paste(code, subcode),
    npt = project %in% c("Admin", "Leave"),
  ) %>% 
  arrange(month, project2, code2) %>% 
  dplyr::filter(!npt) %>% 
  dplyr::select(year, month, project2, code2, hours, budgeted, length) %>% 
  complete(project2, month) %>% # fill gaps
  group_by(month) %>% 
  mutate(
    length = mean(length, na.rm = TRUE), # hours per day
    year = if_else(month(month) >= 6, year(month), year(month) - 1), 
  ) %>% 
  group_by(year, project2) %>% 
  mutate(
    code2 = max(code2, na.rm = TRUE), # fill gaps
  ) %>% 
  group_by(month, length, project2, code2) %>% 
  summarise(
    budgeted = max(0, median(budgeted, na.rm = TRUE), na.rm = TRUE),
    hours = max(0, sum(hours, na.rm = TRUE)),
    profile = round(median(length) * 119 / 8), # target budget per month
    .groups = "keep"
  ) %>% 
  group_by(month) %>% 
  mutate(
    project = str_extract(project2, "[A-Za-z]+"),
    budgeted2 = round(pmax(0, budgeted * profile / sum(budgeted), na.rm = TRUE),1)
  ) %>% 
  ungroup()
  
ggplot(monthly) +
  theme_grey() +
  labs(fill = "Project hours", colour = "", x = "Month", y = "Hours") +
  geom_col(aes(x = month, y = hours, fill = project), position = "dodge") +
  geom_col(aes(x = month, y = budgeted, fill = project), alpha = 0, colour = "black", position = "dodge", linetype = 2) +
  geom_col(aes(x = month, y = budgeted2, fill = project), alpha = 0, colour = "black", position = "dodge") +
  scale_x_date(date_breaks = "1 months", date_labels = "%e %b") +
  guides(colour = "none") +
  theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1))

print(last_plot())

