# README: Sending W5133 emails from a Google Sheet using R

This document describes how to use the email utility to connect to the googlesheet member list and automate the email process while preserving automatic unsubscribe. The first part is a quickstart for R users, followed by detail instructions.

## Quick Start (for R users)

This guide explains how to run `W5133_email.R` to send emails to W5133 members using a Google Sheet and Outlook.

The script:

1. Reads member data from a **Google Sheet**
2. Filters **subscribed members**
3. Builds **personalized unsubscribe links**
4. Creates **Outlook emails**
5. Either **creates drafts** or **sends messages**

---

### 1. Install required packages

Run once:

```r
install.packages(c(
  "googlesheets4",
  "dplyr",
  "glue",
  "Microsoft365R",
  "blastula",
  "htmltools"
))
```

Load the script in RStudio.

---

### 2. Review the SETTINGS section

Near the top of the script.

### Safe test configuration (recommended first run)

```r
test_mode <- TRUE
send_now <- FALSE
draft_limit <- 5
```

Meaning:

| Setting       | Behavior                                   |
| ------------- | ------------------------------------------ |
| `test_mode`   | redirect emails to your test address       |
| `send_now`    | if FALSE, create drafts instead of sending |
| `draft_limit` | limits number of test messages             |

---

### 3. Verify sheet settings

Confirm these variables:

```r
sheet_url <- "https://docs.google.com/spreadsheets/..."
sheet_tab <- "members"
```
The googlesheet is called members in the Wx133.org@gmail.com account.

The sheet should contain at least:

```
email
status
token
member
```

Recipients must satisfy:

```
status == "subscribed"
member == TRUE
token != NA
```

---

### 4. Authenticate

The script will prompt for login the first time.

```r
gs4_auth(email = "wx133.org@gmail.com")
outl <- get_business_outlook()
```

You may see:

* Google authentication (for Sheets)
* Microsoft authentication (for Outlook)

---

### 5. Run the script

In RStudio:

```
Source → W5133_email.R
```

The script will print a **dry-run recipient table** before sending.

Check that the recipients look correct.

---

### 6. Check draft emails

If:

```r
send_now <- FALSE
```

Emails will appear as **drafts in Outlook**.

Review:

* formatting
* links
* unsubscribe URLs
* CC address

---

### 7. Send a single test

Change:

```r
test_mode <- TRUE
send_now <- TRUE
draft_limit <- 1
```

Run again and confirm the email arrives correctly.

---

### 8. Send to all recipients

Only when everything looks correct:

```r
test_mode <- FALSE
send_now <- TRUE
```

Run the script.

Emails will be sent to all subscribed members.






---

## What this script does

This script reads a list of members from a Google Sheet and then creates Outlook emails for the people who should receive a message.

Depending on the settings, it can either:

* **send the emails immediately**, or
* **prepare drafts first** so you can review them

It also creates a personalized **unsubscribe link** for each recipient.

The script is designed for the W5133 mailing workflow. 

---

## What you need before you start

You will need:

* **R** installed on your computer
* Ideally **RStudio** installed as well
* Access to the Google Sheet containing the member list
* Access to the Microsoft 365 / Outlook account that will send the emails
* Internet access, because the script connects to Google Sheets and Outlook

If you are new to R, the easiest way to run this is in **RStudio**.

---

## Files

You should have:

* `W5133_email.R` — the main script
* this README file

---

## What the script expects in the Google Sheet

The script reads a worksheet tab called `"members"` from the Google Sheet. 

The sheet should contain columns with at least these fields:

* `email` — recipient email address
* `status` — should be `"subscribed"` for people who should receive mail
* `token` — used to build the personalized unsubscribe link
* `member` — logical field indicating whether the person should be included

The script cleans these fields a little by trimming spaces and converting `status` to lower case. 

---

## Important warning before you run it

In the current script:

```r
send_now <- TRUE
```

That means the script is currently set to **send real emails immediately** once it gets to the sending loop. 

For safety, change that first to:

```r
send_now <- FALSE
```

That way you can test without sending.

Also consider turning on:

```r
test_mode <- TRUE
```

This will redirect test emails to your own email address instead of real recipients. 

---

## Step 1: Install R and RStudio

If you do not already have them:

1. Install **R**
2. Install **RStudio**

Then open RStudio.

---

## Step 2: Install the required R packages

The script uses several packages. Run this in the R console once:

```r
install.packages(c(
  "googlesheets4",
  "dplyr",
  "glue",
  "Microsoft365R",
  "blastula",
  "htmltools"
))
```

This installs the tools the script needs.

What these packages do, in plain language:

* `googlesheets4` — reads the Google Sheet
* `dplyr` — helps filter and organize the member list
* `glue` — helps build custom text
* `Microsoft365R` — connects to Outlook / Microsoft 365
* `blastula` — helps format emails
* `htmltools` — helps render the email body as HTML

---

## Step 3: Open the script

Open `W5133_email.R` in RStudio.

You will see several sections. The most important one to edit first is the **SETTINGS** section near the top. 

---

## Step 4: Review the settings

Here are the main settings in the script and what they mean.

### Testing options

```r
test_mode <- FALSE
test_email <- "jbayham@colostate.edu"
send_now <- TRUE
draft_limit <- 5
```

What they do:

* `test_mode`

  * `TRUE` = do not send to real members; redirect to `test_email`
  * `FALSE` = use the real addresses from the sheet

* `test_email`

  * where test emails go when `test_mode <- TRUE`

* `send_now`

  * `TRUE` = actually send messages
  * `FALSE` = create drafts only

* `draft_limit`

  * limits how many test emails are created when in test mode

### Recommended safe test setup

Before doing anything else, use:

```r
test_mode <- TRUE
send_now <- FALSE
draft_limit <- 5
```

This is the safest combination for testing.

---

### Sheet settings

```r
sheet_url <- "https://docs.google.com/spreadsheets/d/..."
sheet_tab <- "members"
```

* `sheet_url` tells the script where the member list lives
* `sheet_tab` tells it which worksheet tab to read

---

### Unsubscribe link base

```r
base_unsub_url <- "https://script.google.com/macros/s/..."
```

This is the web address used to build the unsubscribe link for each member. The script appends each person’s token to this URL. 

---

### Email settings

```r
email_subject <- "W5133 Annual Reporting"
cc_email = "wx133.org@gmail.com"
```

* `email_subject` is the subject line
* `cc_email` is the address copied on each message

---

## Step 5: Authenticate the first time you run it

The script includes these authentication steps:

```r
gs4_auth(email = "wx133.org@gmail.com")
outl <- get_business_outlook()
```

The first time you run the script, R will usually open a browser window and ask you to sign in.

You may need to:

* sign into Google for Sheet access
* sign into Microsoft 365 for Outlook access
* grant permission for R to use those services

This is normal.

---

## Step 6: Run the script safely

If you are new to R, the easiest way is:

1. Open the script in RStudio
2. Change the settings to safe test values:

   ```r
   test_mode <- TRUE
   send_now <- FALSE
   ```
3. Click **Source** in RStudio

Or highlight all the code and run it.

---

## What the script does when it runs

The script does the following:

### 1. Reads the member list from Google Sheets

It loads the member worksheet and standardizes some column names and values. 

### 2. Filters the recipient list

It keeps only rows where:

* `email` is present
* `status == "subscribed"`
* `token` is present
* `member` is true

That means only subscribed members with a valid unsubscribe token are included. 

### 3. Creates unsubscribe URLs

Each recipient gets their own unsubscribe link based on their token. 

### 4. Prints a dry-run table

The script prints a table showing who would receive an email. This is very useful for checking before sending. 

### 5. Builds the HTML email body

The body includes:

* a thank-you note about the Boise meeting
* a link to the Google reporting form
* a note about why they are receiving the email
* an unsubscribe link
* the project website
* the note about trusting `wx133.org@gmail.com`

### 6. Creates Outlook email objects

For each recipient, the script creates a message in Outlook. Depending on settings, it either:

* sends it immediately, or
* leaves it as a draft

---

## Suggested first test

Use this exact setup first:

```r
test_mode <- TRUE
test_email <- "your_email@yourdomain.edu"
send_now <- FALSE
draft_limit <- 3
```

Then run the script.

What to check:

* Did the script authenticate successfully?
* Did it read the Google Sheet?
* Did the dry-run table look correct?
* Did the draft messages appear in Outlook?
* Did the unsubscribe links look correct?
* Did the HTML formatting render properly?

Only after that should you try sending.

---

## Suggested second test

Once drafts look correct, try:

```r
test_mode <- TRUE
send_now <- TRUE
draft_limit <- 1
```

This sends one real test message to yourself.

Check:

* Did it arrive?
* Did it go to spam?
* Did the links work?
* Did the CC behave as expected?

---

## Final production run

Only when you are confident everything is correct:

```r
test_mode <- FALSE
send_now <- TRUE
```

Be careful: this will send to the real recipient list pulled from the sheet. 

---

## Good safety practices

Before a real run:

* make sure `status` values in the sheet are correct
* make sure tokens are present
* confirm the `member` column is correct
* test the unsubscribe link
* send a test to yourself first
* verify the Google Form link is the correct one
* verify the CC address is intended

A very good extra precaution is to temporarily uncomment this line in the script:

```r
#stop("Dry-run complete. Review recipients above, then comment out this line to create drafts.")
```

If you remove the `#`, the script will stop after printing the dry-run table, before creating any messages. That is a useful safety checkpoint. 

---

## Common problems and fixes

### Problem: “package not found”

Install the packages again:

```r
install.packages(c(
  "googlesheets4",
  "dplyr",
  "glue",
  "Microsoft365R",
  "blastula",
  "htmltools"
))
```

---

### Problem: Google authentication fails

Try:

* making sure you are signed into the correct Google account
* rerunning the script
* restarting RStudio and trying again

---

### Problem: Outlook authentication fails

Make sure:

* you are using the correct Microsoft 365 account
* your organization allows API access
* any browser sign-in popup is completed

---

### Problem: No recipients found

The script will stop with:

```r
stop("No eligible subscribed recipients with tokens were found.")
```

This usually means one of these is true:

* `email` is blank
* `status` is not `"subscribed"`
* `token` is blank
* `member` is false or missing
* the script is pointing at the wrong sheet/tab

---

### Problem: Emails send immediately when I only wanted drafts

Check this line:

```r
send_now <- TRUE
```

Change it to:

```r
send_now <- FALSE
```

---

### Problem: Test mode still seems wrong

In test mode, the script rewrites the recipient email to `test_email`, but keeps the original address in `actual_recipient` for reference. That behavior is intentional. 

---

## How to edit the email text

The message text lives inside this function:

```r
build_email_html <- function(unsub_url) {
  ...
}
```

You can edit the wording there.

A few tips:

* Keep the unsubscribe link in place
* Be careful with quotes and HTML tags like `<p>` and `<a>`
* After editing, always send a test to yourself first

---

## Plain-language summary of the workflow

Here is the full process in simple terms:

1. R connects to your Google Sheet
2. It reads the member list
3. It keeps only eligible subscribed members
4. It builds one custom unsubscribe link per person
5. It creates one email per person in Outlook
6. It either drafts or sends those messages depending on your settings

---

## Recommended beginner workflow

Each time you use the script:

1. Open `W5133_email.R`
2. Set:

   ```r
   test_mode <- TRUE
   send_now <- FALSE
   ```
3. Run the script
4. Check the printed recipient table
5. Check the Outlook drafts
6. Send one test to yourself
7. Only then switch to the real run

---

## Optional improvements you may want later

As you get more comfortable with R, you may want to improve the script by adding:

* logging to a CSV file
* better error handling if one email fails
* a stronger “draft only” mode
* a separate preview mode
* duplicate email detection
* validation checks on the Google Sheet before sending

---

## Contact notes for future users

This script uses both Google and Microsoft authentication, so the first run can feel awkward if you are not used to R. That is normal. Start with test mode, keep `send_now <- FALSE`, and verify every step before sending to the full list.
