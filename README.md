# wos-tools_issue84

## Overview
This project consists of the following set of scripts:
- Script to extract Gmail threads based on specified conditions and save them as Google Documents on Google Drive.
- Script that extracts the title and URL of a Google document and outputs it to a spreadsheet.
- Script that sends the URL of the spreadsheet to Google Space.

## Purpose
The purpose of this project is to manage the progress of queries to wos-tools.

## Configuration
- The script is written in Google Apps Script and is designed to be run as a time-driven trigger, specifically a "date-based timer" trigger.
- Script properties are used to store configuration information, including the output folder ID, sender's email address, Google Doc template ID, and webhook URL (for sending the spreadsheet URL to a Google Space).

## Usage
1. Set up a Google Apps Script project.
2. Configure the script properties with the necessary information:
   - `outputFolderId`: The ID of the Google Drive folder where Google Documents will be saved.
   - `senderEmail`: The sender's email address to filter Gmail threads.
   - `templateDocId`: The ID of the Google Doc template.
   - `webhookUrl`: The webhook URL for sending the spreadsheet URL to a Google Space.
3. Create a time-driven trigger in the Google Apps Script project, such as a "date-based timer" trigger.
4. At the timing specified in the trigger, this script automatically searches for Gmail threads that match the specified criteria and saves them as Google Documents in the folder specified in the properties.
5. Send a link to the spreadsheet to Google Space at the time specified in the trigger.
