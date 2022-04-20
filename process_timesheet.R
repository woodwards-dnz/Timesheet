
# timesheet analysis 

suppressWarnings({
library(readxl)
library(janitor)
library(dplyr)
library(tidyr)
library(lubridate)
library(ggplot2)
})

path <- "Timesheet2022.xlsx"
done <- ymd("2022-04-11") # monday
wdays <- wday(done + days(0:6), week_start = 1, label = TRUE, abbr = TRUE)
thisyear <- "2021-2022" # select projects from this year

charges <- read_excel(path, sheet = "Charges", skip = 1) %>% clean_names() %>% drop_na(date)

charges %>% 
  group_by(period) %>% 
  summarise(
    weekhours = sum(hours)
  ) %>% 
  ungroup() %>% 
  mutate(
    balance = cumsum(weekhours) - seq(n()) * 36  
  ) %>% 
  ggplot() +
  labs(fill = "", colour = "") +
  geom_col(aes(x = period, y = weekhours, fill = "Week hours")) +
  geom_hline(yintercept = 36, linetype = 2) +
  geom_line(aes(x = period, y = balance, colour = "Balance")) +
  geom_point(aes(x = period, y = balance, colour = "Balance", fill = "Balance")) 

print(last_plot())

projects <- read_excel(path, sheet = "Projects", skip = 1) %>% clean_names()

projects <- projects %>% 
  mutate(year = ifelse(is.na(year), thisyear, year))

timesheet <- charges %>%
  left_join(projects, by = c("year", "project", "subproject")) %>% 
  arrange(year, period, date, project, subproject) %>% 
  mutate(
    date = wday(date, week_start = 1, label = TRUE, abbr = TRUE)
  ) %>% # use day of week
  group_by(year, period, date, project, subproject, code, subcode) %>% 
  summarise(
    hours = sum(hours),
    .groups = "keep"
  ) %>% 
  pivot_wider(names_from = "date", values_from = "hours") %>% 
  mutate(
    Total = rowSums(across(any_of(wdays)), na.rm = TRUE) 
  ) %>% 
  group_by(year, period) %>% 
  mutate(
    period = as_date(period),
    Week = sum(Total, na.rm = TRUE)
  ) %>% 
  ungroup() %>% 
  dplyr::filter(period > done) %>% 
  dplyr::select(-year, -subproject, -subcode)

print(timesheet)
  