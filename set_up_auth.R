
gitcreds::gitcreds_set() # provide PAT token

usethis::git_remotes() # check https or ssh

usethis::use_git_remote( # use https
  "origin",
  "https://github.com/woodwards/Timesheet.git",
  overwrite = TRUE
)

