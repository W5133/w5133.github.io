#Read member list from googlesheet and send out


# Packages
library(googlesheets4)
library(dplyr)
library(glue)
library(Microsoft365R)
library(blastula)
library(htmltools)

# -----------------------------
# 1) SETTINGS
# -----------------------------

# Optional: send drafts only to yourself while testing
test_mode <- FALSE
test_email <- "jbayham@colostate.edu"

#Send emails or just draft (FALSE)
send_now <- TRUE

# Google Sheet URL or Sheet ID
sheet_url <- "https://docs.google.com/spreadsheets/d/1JEyDDqVjKpPppmDvRs0W17pux8bbUMBEzPLHNS_rH-8"

# Name of the worksheet tab
sheet_tab <- "members"

# Your Apps Script unsubscribe endpoint
base_unsub_url <- "https://script.google.com/macros/s/AKfycbzeRtpmrqY5dwN6l-KMyA-5odAi0Bl5jjSGpeiu4KzxJK8RsOYa-SJuH43b1MAu37fjhQ/exec"

# Optional: cap number of drafts created while testing
draft_limit <- 5

# Subject line
email_subject <- "W5133 Annual Reporting"

cc_email = "wx133.org@gmail.com"

# External HTML email template file
#email_template_file <- "e1_post_meeting.txt"
email_template_file <- "e2_reminder1.txt"

# -----------------------------
# 2) AUTHENTICATE
# -----------------------------

# Google Sheets auth
# First run will usually open a browser for OAuth
gs4_auth(email = "wx133.org@gmail.com")

# Outlook / Microsoft 365 auth
# First run will usually open a browser for OAuth
outl <- get_business_outlook()

# -----------------------------
# 3) READ MEMBER LIST
# -----------------------------

members <- read_sheet(sheet_url, sheet = sheet_tab) %>%
  rename_with(tolower) %>%
  mutate(
    row_id = row_number() + 1,   # +1 because sheet row 1 is header
    email = trimws(email),
    cc_email = cc_email,
    status = tolower(trimws(status)),
    token = trimws(token)
  )

# Keep only active subscribed members with both email and token
recipients <- members %>%
  filter(
    !is.na(email), email != "",
    status == "subscribed",
    is.na(report_submitted), 
    !is.na(token), token != ""
  ) %>%
  filter(member) %>%
  mutate(
    unsubscribe_url = paste0(base_unsub_url, "?u=", utils::URLencode(token, reserved = TRUE))
  )

# Optional testing behavior
if (test_mode) {
  recipients <- recipients %>%
    slice_head(n = draft_limit) %>%
    mutate(actual_recipient = email,
           email = test_email)
} else {
  recipients <- recipients %>%
    mutate(actual_recipient = email)
}

# Safety check
if (nrow(recipients) == 0) {
  stop("No eligible subscribed recipients with tokens were found.")
}

# -----------------------------
# 4) DRY-RUN TABLE
# -----------------------------

dry_run <- recipients %>%
  transmute(
    row_id,
    actual_recipient,
    draft_to = email,
    status,
    token,
    unsubscribe_url
  )

print(dry_run, n = nrow(dry_run))

# Optional safety stop:
# Uncomment this if you always want to inspect first and then rerun
#stop("Dry-run complete. Review recipients above, then comment out this line to create drafts.")



# -----------------------------
# 5) EMAIL BODY FUNCTION
# -----------------------------


build_email_html <- function(unsub_url) {
  candidate_paths <- c(
    email_template_file,
    file.path("email_management", email_template_file)
  )
  template_path <- candidate_paths[file.exists(candidate_paths)][1]

  if (is.na(template_path)) {
    stop("Template file not found. Checked: ", paste(candidate_paths, collapse = ", "))
  }

  template_html <- paste(readLines(template_path, warn = FALSE), collapse = "\n")
  glue(template_html)
}

# -----------------------------
# 5) CREATE DRAFTS IN OUTLOOK
# -----------------------------

draft_results <- vector("list", nrow(recipients))

for (i in seq_len(nrow(recipients))) {
  to_email <- recipients$email[i]
  unsub_url <- recipients$unsubscribe_url[i]
  
  html_body <- build_email_html(unsub_url)
  
  email_obj <- compose_email(
    body = htmltools::HTML(html_body)
  )
  
  #Construct message
  msg <- outl$create_email(
    email_obj,
    subject = email_subject,
    to = to_email,
    cc = cc_email
  )
  
  #Toggle above to send
  if (send_now) {
    msg$send()
    action <- "sent"
    Sys.sleep(10)
    message("Sent to ",to_email)
  } else {
    action <- "drafted"
  }
  
  draft_results[[i]] <- tibble(
    row_id = recipients$row_id[i],
    actual_recipient = recipients$actual_recipient[i],
    draft_to = to_email,
    action = action,
    action_time = as.character(Sys.time())
  )
}

message("Finished")
