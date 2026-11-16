# Porting the Email Feature to Another Project

This guide explains how to take the email generation system built for the
PM Requests Report Builder and adapt it for another Excel/VBA project.

## What the feature does

- Stores an email template (HTML with formatting) in a hidden sheet (EmailTemplate)
- Stores CC/To/ImagePath settings in a hidden sheet (EmailSettings)
- Adds buttons on START HERE: Edit CC Emails, Edit To Emails, Generate Email
- Each data tab gets a Generate Email button
- Creates Outlook email with subject (today date), CC/To, HTML body, data table, image
- Sheets hidden by default; Edit buttons unhide; Done button re-hides

## VBA procedures to copy

1. GetEmailTemplate() - reads HTML template from EmailTemplate sheet
2. GetCCEmails() - reads CC from EmailSettings B1
3. GetToEmails() - reads To from EmailSettings B3
4. EditCCEmails() - unhides EmailSettings, activates B1
5. EditToEmails() - unhides EmailSettings, activates B3
6. DoneEditing() - hides active sheet, returns to START HERE
7. GenerateManagerEmail() - creates email from active sheet data
8. CreateEmailFromReport(outWb) - creates email from workbook data
9. SetupSheetVisibility() - moves START HERE first, hides all others
10. SetupEmailButtons() - creates Edit CC / Edit To buttons
11. SetupGenerateEmailButton() - creates Generate Email button

## Sheets to create

EmailTemplate (hidden): A1=HTML template with placeholder, C1-C3=instructions
EmailSettings (hidden): A1=CC Emails, B1=value, C1=instructions; A2=Image Path, B2=value; A3=To Emails, B3=value; A5=title, A6=help

## Buttons on START HERE

- SetupSheetVisibility (hide all tabs except START HERE)
- SetupEmailButtons (Edit CC Emails, Edit To Emails at rows 23-24)
- SetupGenerateEmailButton (Generate Email)

## Button on data tabs

AddShape at column K, row 1: green rounded rect, text Generate Email, OnAction GenerateManagerEmail

## Adaptation steps

1. Copy VBA procedures into your module
2. Create EmailTemplate and EmailSettings sheets
3. Add buttons to START HERE
4. Add Generate Email button to data tabs
5. Convert email template to HTML (bold, bullets, highlights)
6. Store HTML in EmailTemplate A1 with placeholder
7. Test

## Key design decisions

- HTML body preserves formatting (bold, highlights, bullets)
- Data as HTML table (single-line cells)
- Hidden sheets (clean workbook, buttons unhide/rehide)
- Image at bottom (attached with Content-ID)
- No InputBox (edit sheets directly)
- Done button (returns to START HERE, re-hides)

## Dependencies

- Outlook (late binding, no references needed)
- No external libraries
- Windows only (Mac may need adjustments)
