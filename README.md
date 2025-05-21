# Email Automation System

This Excel macro automates the process of sending personalized emails to different users with specific attachments for each recipient.

## Overview

The system reads data from an Excel table and sends customized emails to store managers and area managers, with specific Excel and PDF reports attached for each store.

## Features

- Automated email sending to multiple recipients
- Customized email content for each store
- Automatic attachment of store-specific reports
- Support for both Excel (.xlsx) and PDF files
- Error handling for missing files

## Requirements

- Microsoft Excel
- Microsoft Outlook
- A data folder containing the reports

## Setup

1. Create a folder named `data` in the same directory as your Excel workbook
2. Place your store reports in the `data` folder with the following naming convention:
   - Excel files: `[storeID].xlsx`
   - PDF files: `[storeID].pdf`

## Excel Table Structure

The macro reads from a table named "taula_emails" with the following columns:
1. Store Name
2. Store Manager Email
3. Area Manager Email
4. Coordinator Email
5. Office Email
6. Project Manager Email
7. File Path (store ID for attachments)

## Email Content

Each email includes:
- Personalized greeting
- Store-specific information
- Both Excel and PDF attachments
- Professional signature

## Error Handling

The system includes error handling for:
- Missing attachment files
- Invalid email addresses
- File path issues

## Usage

1. Open the Excel workbook
2. Ensure the "taula_emails" table is properly filled with recipient information
3. Place the corresponding report files in the data folder
4. Run the macro `SendCustomStoreEmails`

## Notes

- The macro automatically uses the previous month's name and year in the email subject and content
- Emails are only sent if both Excel and PDF files exist for the store
- The system supports multiple recipients in To and CC fields
