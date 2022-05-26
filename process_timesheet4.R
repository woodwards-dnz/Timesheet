
# timesheet analysis ####

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


stop()

# monthly project totals ####
monthly <- combined %>%
  mutate(
    project2 = paste(project, subproject),
    code2 = paste(code, subcode),
  ) %>% 
  dplyr::select(year, month, project2, code2, hours, budgeted, length) %>% 
  complete(project2, month) %>% 
  group_by(month) %>% 
  mutate(
    length = mean(length, na.rm = TRUE),
    year = if_else(month(month) >= 6, year(month), year(month) - 1), 
  ) %>% 
  group_by(year, project2) %>% 
  mutate(
    code2 = max(code2, na.rm = TRUE),
  ) %>% 
  group_by(month, length, project2, code2) %>% 
  summarise(
    budgeted = pmax(0, median(budgeted, na.rm = TRUE), na.rm = TRUE),
    hours = pmax(0, sum(hours, na.rm = TRUE)),
    profile = length * 114 / 8,
  ) %>% 
  arrange(project2, code2, month) 
  
  
    


  
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

