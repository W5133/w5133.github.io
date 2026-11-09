# Sending W5133 emails with R

This folder contains an Outlook mailer and a separate CSV cleaning utility. The CSV cleaner is not needed for each mailing.

## Files

| File | Purpose |
| --- | --- |
| `W5133_email.R` | Reads eligible recipients from Google Sheets and creates Outlook drafts or sends individual messages with personalized Google Form unsubscribe links. |
| `w5133_email_clean.R` | Cleans and checks local CSV email lists, flags potential problems, and joins a separate member list to produce a CSV with a membership flag. Does not send email or update Google Sheets. |
| `e1_post_meeting.txt` | HTML template for the post-meeting reporting request. |
| `e2_reminder1.txt` | HTML template for the reporting reminder. |
| `e3_2027_meeting_announcement.txt` | HTML template for the 2027 meeting announcement; currently selected by the sending script. |

## Quick start: prepare and send an email

### 1. Open the project and install packages if needed

Open `w5133.github.io.Rproj` in RStudio. Run this once if the packages are not already installed:

```r
install.packages(c(
  "googlesheets4", "dplyr", "glue",
  "Microsoft365R", "blastula", "htmltools"
))
```

Open `email_management/W5133_email.R`.

### 2. Set up a draft test before running

The checked-in script currently has `test_mode <- FALSE` and `send_now <- TRUE`, which sends to real recipients. Change the settings first:

```r
test_mode <- TRUE
test_email <- "jbayham@colostate.edu"
send_now <- FALSE
draft_limit <- 1
```

| Setting | Behavior |
| --- | --- |
| `test_mode <- TRUE` | Uses the first `draft_limit` eligible rows and redirects their To addresses to `test_email`. |
| `test_mode <- FALSE` | Uses all eligible rows and their real addresses. |
| `send_now <- FALSE` | Creates Outlook drafts. |
| `send_now <- TRUE` | Sends messages, with a 10-second pause after each send. |
| `draft_limit` | Limits messages only in test mode, whether drafting or sending. |
| `cc_email` | CC address on every message, including tests; currently `wx133.org@gmail.com`. |

Test mode redirects the To address and prefills the test email address in the unsubscribe form. The CC remains active. Opening the form does not submit an unsubscribe request.

### 3. Edit the subject and template

Set the subject and choose the template near the top of the script:

```r
email_subject <- "W5133 Annual Reporting"
email_template_file <- "e2_reminder1.txt"
```

Edit the email wording in the selected `.txt` file. These files contain HTML, despite their extension. Update dates and the Google Form link for the new mailing, and preserve the `{unsub_url}` placeholder in unsubscribe links.

The script searches for the template in the current working directory and then in `email_management/`, so it works from either the project root or that folder. The `build_email_html()` function reads the template and substitutes the recipient's unsubscribe URL.

### 4. Check the member sheet and recipient filter

The configured [Google Sheet](https://docs.google.com/spreadsheets/d/1JEyDDqVjKpPppmDvRs0W17pux8bbUMBEzPLHNS_rH-8) belongs to the `wx133.org@gmail.com` workflow. The script reads its `members` tab and requires these columns:

| Column | Requirement for inclusion |
| --- | --- |
| `email` | Nonmissing and nonempty after trimming whitespace. |
| `status` | `subscribed`, after trimming whitespace and converting to lowercase. |
| `token` | No longer required or used. Existing values can remain in the sheet. |
| `member` | Used only when `member_filter` is not `"all"`; logical `TRUE` for members or `FALSE` for nonmembers. Missing values are excluded when filtering. |
| `report_submitted` | Used only when `report_filter` is not `"all"`; missing (`NA`) means not submitted. |

Choose the audience in SETTINGS at the top of `W5133_email.R`; no edits to the filtering code are needed:

```r
member_filter <- "members"         # "all", "members", or "nonmembers"
report_filter <- "not_submitted"   # "all", "not_submitted", or "submitted"
```

The example above selects the reporting-reminder audience. The current script selects members regardless of reporting (`member_filter <- "members"`, `report_filter <- "all"`). `"submitted"` means any nonmissing value, including `FALSE` or text, matching the existing sheet convention. For an announcement to everyone subscribed, set both filters to `"all"`. For all subscribed members regardless of reporting, use `member_filter <- "members"` and `report_filter <- "all"`. Subscription and email requirements always apply.

The script does not deduplicate addresses. Each eligible row produces a message.

### 5. Run and authenticate

From the project root, run:

```r
source("email_management/W5133_email.R")
```

Or use Source in RStudio with the project root as the working directory.

The script authenticates using:

```r
gs4_auth(email = "wx133.org@gmail.com")
outl <- get_business_outlook()
```

Complete any Google and Microsoft browser sign-in prompts. Use the Microsoft 365 account you intend to send from; the Google account provides Sheet access, not the sending mailbox.

The script prints a recipient table, then immediately proceeds to creating drafts or sending messages. **The table is not an interactive confirmation or a pause.** To stop after inspecting recipients, uncomment this existing line before running:

```r
stop("Dry-run complete. Review recipients above, then comment out this line to create drafts.")
```

Comment it out again when ready to create messages.

### 6. Review the Outlook draft and send one test

With `send_now <- FALSE`, inspect Outlook Drafts for the subject, formatting, dates, form link, To and CC addresses, and unsubscribe URL. Open the unsubscribe link and verify the prefilled test email address; leave the confirmation unchecked and do not submit a request for a real member.

To send one test message:

```r
test_mode <- TRUE
send_now <- TRUE
draft_limit <- 1
```

Run the script again and check the delivered message. This creates a new message; it does not send the earlier draft.

### 7. Send the mailing

After reviewing the content and recipient filter, set:

```r
test_mode <- FALSE
send_now <- TRUE
```

Run the script. It sends one message per eligible row and CCs `cc_email` on each message. To prepare drafts addressed to all eligible recipients instead, use `test_mode <- FALSE` and `send_now <- FALSE`.

There is no persistent sent-message log or automatic resume/deduplication. Results are collected in the R object `draft_results`, but are not saved to a file. Rerunning the script can send duplicate messages, including after a partially completed run. Check Outlook Sent Items before retrying a failed mailing.

## The separate email-cleaning utility

`w5133_email_clean.R` is for preparing local mailing-list data. It is independent of the Google Sheet/Outlook sending workflow.

The script currently changes the working directory to `~/Documents` and uses these files:

| File | Role |
| --- | --- |
| `w5133_emails_2026.csv` | Input address list, with an `email` column or separate local-part and domain columns. |
| `w5133_member_emails.csv` | Separate member list with an `email` column, used to assign membership. |
| `w5133_emails_validated.csv` | Output with normalized addresses, validation results, and review flags. |
| `w5133_emails_cleaned.csv` | Output combining unique cleaned addresses with the member list and a `member` flag. |

Before running, review the working directory and filenames. Existing output files at those paths will be overwritten. Install its packages if needed:

```r
install.packages(c("readr", "dplyr", "stringr", "purrr", "tibble", "curl"))
```

The utility:

1. Extracts and normalizes addresses, or reconstructs them from local-part/domain columns.
2. Checks syntax and flags missing addresses, obvious junk, duplicates, role accounts, and disposable domains.
3. Attempts DNS checks for domain resolution and MX records, and assigns a validation status.
4. Writes the validation CSV.
5. Keeps unique, syntactically valid addresses from the main input, then full-joins the separate member list and writes the cleaned CSV with membership flags.

The cleaned output is **not restricted to addresses classified as `likely_valid`**: the main-input filter checks only presence and syntax, and the full join can add addresses from the separate member list. Review the validation output before using the result. DNS checks do not prove an individual mailbox exists.

The cleaning script does not create subscription statuses or unsubscribe tokens, upload its output to Google Sheets, or send email. Managing those fields in the member sheet is a separate step.

## Troubleshooting

- **Package not found:** Install the packages listed for the script you are using.
- **Google authentication fails:** Confirm access to the configured sheet as `wx133.org@gmail.com` and retry authentication.
- **Outlook authentication fails:** Confirm the intended Microsoft 365 account and that organizational settings allow API access.
- **No eligible recipients:** Check `email`, `status`, `member`, and the current `report_submitted` filter.
- **Missing `report_submitted` or `member` column:** Add the column if filtering on it, or set the corresponding filter to `"all"`.
- **Template file not found:** Run from the project root or `email_management/`, and verify `email_template_file` matches an existing file.
- **More drafts than expected:** `draft_limit` applies only when `test_mode <- TRUE`; each rerun creates additional messages.

## Google Form unsubscribe workflow

The published unsubscribe form is configured by `base_unsub_url` and `unsub_email_entry` in the mailer. Each footer opens the form with the destination email address URL-encoded in `entry.2066026793`. Reasons and confirmation are not prefilled; the recipient must submit the form. The HTML renderer escapes the URL for use in an HTML attribute.

Responses populate **Form Responses 3** in the existing member-list spreadsheet. Review these requests and update member subscription statuses before each mailing. The mailer does not process requests automatically.

Previously sent emails still contain the old Apps Script token links. On November 8, 2026, the original endpoint was updated in place to version 6 of `member_management`. Its `doGet` handler now ignores tokens and displays a link to the new form without reading or changing member records. People following an old link must enter their email and submit the form. The new R email links prefill the email directly.

