
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

# close Excel ####
print("WARNING --- closing Excel!!!")
system("taskkill /IM Excel.exe")
Sys.sleep(1)

# options ####
options(dplyr.summarise.inform = FALSE)
path <- "TimesheetB2023.xlsx"
done <- ymd("2023-07-31") # monday
print(paste("Done to", done))
wdays <- wday(done + days(0:6), week_start = 1, label = TRUE, abbr = TRUE)

# functions ####
zero2na <- function(x){
  ifelse(x <= 0, x * NA, x)
}
na2zero <- function(x){
  ifelse(is.na(x), 0, x)
}

# read timesheet ####
charges <- read_excel(path, sheet = "Charges", skip = 1) %>% clean_names() %>% drop_na(date, hours) %>% 
  mutate(
    date = as_date(date),
    period = as_date(period),
    date = pmax(date, period), # merge Feb 2022 dates into March 2022
    month = dmy(paste("1", month(date), year(date))),
    year = if_else(month(date) >= 6, year(date), year(date) - 1),
  ) %>% 
  dplyr::select(year, month, period, date, project, code, hours, length) %>% 
  arrange(date, project, code) %>% 
  group_by(project) %>% 
  fill(code) %>% # fill code down
  ungroup()
stopifnot(all(!is.na(charges$code)))

print(paste("Timesheet to", max(charges$date)))

# complete table ####
projects <- charges %>% 
  group_by(year, project, code) %>% 
  summarise(
    startdate = min(date, na.rm = TRUE),
    enddate = max(date, na.rm = TRUE)
  ) %>% 
  ungroup()
blanks <- charges %>% 
  dplyr::select(year, month, period) %>% 
  distinct() %>% 
  mutate(date = period) %>% 
  left_join(projects, by = "year", multiple = "all", relationship = "many-to-many") %>% 
  mutate(hours = 0, length = NA_real_) %>% 
  dplyr::select(-startdate, -enddate)
combined <- charges %>% 
  bind_rows(blanks) %>% 
  left_join(projects, by = c("year", "project", "code")) %>% 
  dplyr::filter(date >= startdate & date <= enddate) 

# weekly and running totals ####
print("Plot Weekly and Running Totals:")

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
  labs(title = "Weekly and Running Totals", fill = "", colour = "", x = "Week", y = "Hours") +
  geom_col(aes(x = period, y = weekhours, fill = "Week Hours")) +
  geom_hline(yintercept = seq(7.2,36,7.2), linetype = 2) +
  geom_line(aes(x = period, y = balance, colour = "Balance"), linewidth = 1.5) +
  geom_point(aes(x = period, y = balance, colour = "Balance", fill = "Balance"), size = 4) +
  scale_x_date(date_breaks = "1 weeks", date_labels = "%e %b") +
  guides(colour = "none") +
  theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
  scale_y_continuous(breaks = seq(7.2,36,7.2))

print(last_plot())


# timesheet ####
print("Timesheet Table:")

timesheet <- combined %>% 
  dplyr::filter(period > done) %>% 
  mutate(
    wday = wday(date, week_start = 1, label = TRUE, abbr = TRUE),
    date = NULL,
  ) %>% 
  arrange(period, wday, project, code) %>% 
  group_by(period, wday, project, code) %>% 
  summarize(
    hours = zero2na(sum(hours)),
    .groups = "keep"
  ) %>% 
  ungroup() %>% 
  pivot_wider(id_cols = all_of(c("period", "project", "code")), names_from = "wday", values_from = "hours") %>% 
  mutate(
    Total = rowSums(across(any_of(wdays)), na.rm = TRUE) 
  ) %>% 
  group_by(period) %>% 
  mutate(
    period = as_date(period),
    Week = sum(Total, na.rm = TRUE)
  ) %>% 
  ungroup() %>% 
  arrange(period, project) %>% 
  dplyr::filter(Total > 0)

print(timesheet)


# monthly project totals ####
print("Plot Monthly Project Budgets:")

NPT <- c("Admin", "Learning", "Leave")
NPTbudgeted <- 174 - 119
  
monthly <- combined %>%
  mutate(
    project2 = ifelse(project %in% NPT, "NPT", project),
    code2 = ifelse(project %in% NPT, "NPT", code),
  ) %>% 
  arrange(month, project2, code2) %>% 
  dplyr::select(year, month, project2, code2, hours, length) %>% 
  complete(project2, month, fill = list(hours = 0)) %>% # fill gaps
  group_by(month) %>% 
  mutate(
    length = mean(length, na.rm = TRUE), # hours per day
    year = if_else(month(month) >= 6, year(month), year(month) - 1), # fill gaps
  ) %>% 
  group_by(year, project2) %>% 
  mutate(
    code2 = if_else(is.na(code2), "0", code2), # fill gaps
    code2 = max(code2, na.rm = TRUE), # fill gaps
  ) %>% 
  arrange(year, month, length, project2, code2) %>% 
  group_by(year, month, length, project2, code2) %>% 
  summarise(
    hours = max(0, sum(hours, na.rm = TRUE)),
    profile = round(median(length) * 119 / 8), # target budget per month
    .groups = "keep"
  ) %>% 
  mutate(
    project = str_extract(project2, "[A-Za-z]+"),
  ) %>% 
  ungroup()
  
ggplot(monthly) +
  theme_grey() +
  labs(title = "Monthly Project Budgets", fill = "Project", colour = "", x = "Month", y = "Hours") +
  geom_col(aes(x = month, y = hours, fill = project), colour = NA, alpha = 0.4, position = "dodge") +
  scale_x_date(date_breaks = "1 months", date_labels = "%e %b") +
  guides(colour = "none") +
  theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1))

print(last_plot())

ggplot(monthly) +
  theme_grey() +
  labs(title = "Monthly Project Budgets", fill = "Project", colour = "", x = "Month", y = "Hours") +
  geom_col(aes(x = project, y = hours, fill = project), colour = NA, alpha = 0.4, position = "dodge") +
  guides(colour = "none") +
  theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
  facet_wrap("month", labeller = function(x) format(x, '%b-%y'))

print(last_plot())


# yearly project totals ####
print("Yearly Project Totals:")

monthly %>% 
  arrange(year, project) %>% 
  group_by(year, project) %>% 
  summarise(
    months = n_distinct(month),
    # profile = median(profile),
    hours = round(sum(hours),0),
    hours_month = round(sum(hours) / months,0),
  ) %>% 
  dplyr::filter(hours > 0)
    
