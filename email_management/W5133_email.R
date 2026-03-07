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
  glue("
<p>Hello W5133 Member or Affiliate,</p>

<p>Thank you to Lee for organizing a great meeting in Boise. I had a good time, and I think everyone else did too.</p>

<p>It is reporting time again. We have 30 days from the conclusion of the annual meeting (March 27) to compile and submit our report. Similar to last year, we will be collecting information via a <a href='https://forms.gle/114uCWJ9Eec9xKhK8'>Google Form</a> to streamline the reporting process. Please complete this form as soon as possible to avoid repeated reminder emails from me. This is particularly important for all project members and those at Land Grant institutions.</p>

<p><strong>Google Form Link:</strong>
<a href='https://forms.gle/114uCWJ9Eec9xKhK8'>https://forms.gle/114uCWJ9Eec9xKhK8</a></p>

<p>If you have any questions, please let me know. I look forward to seeing you all again next year.</p>

<hr>
<p><strong>Note:</strong> If you are receiving this email, you are either on the master list for the W5133 Multistate Research Project or listed as a project participant in NIMSS. If you would like to be removed, click <a href='{unsub_url}'>Unsubscribe</a>.</p>

<p>As a reminder, Lee created this great website for our multistate project,
<a href='https://w5133.github.io/'>https://w5133.github.io/</a>,
where you can find past and upcoming meeting information.</p>

<p>I've cc'ed an email address, Wx133.org@gmail.com, that I would like to use for future correspondence. In tests on myself, these emails always go to spam. Adding this email to your trusted list will aid in future communication.</p>

<hr>
<p style='font-size:12px;color:#666;'>
To unsubscribe from these emails, click here:
<a href='{unsub_url}'>Unsubscribe</a>
</p>
")
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
