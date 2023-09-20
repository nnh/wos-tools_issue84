# wos-tools_issue84
## Overview
This project consists of the following scripts  
- Script to extract Gmail threads based on specified conditions and store them as Google Documents in Google Drive folder
- Script to retrieve information of Google Documents under the specified folder and output them to Google Spreadsheet
- Script to send the URL of the spreadsheet to Google Space
## Purpose
The goal of this project is to manage progress on wos-tools inquiries.
## Configuration
The script is written in Google Apps Script and is designed to be run as a time-driven trigger, specifically a "date-based timer" trigger.
Script properties are used to store configuration information, including the output folder ID, sender's email address, Google Doc template ID, and webhook URL (for sending the spreadsheet URL to a Google Space).
## Usage
1. Set up a Google Apps Script project.  
2. Configure the script properties with the necessary information:  
    - outputFolderId: The ID of the Google Drive folder where Google Document will be saved.  
    - senderEmail: The sender's email address to filter Gmail threads.  
    - templateDocId: The ID of the Google Doc template.  
    - webhookUrl: The webhook URL for sending the spreadsheet URL to a Google Space.  
3. Create a time-driven trigger in the Google Apps Script project, such as a "date-based timer" trigger, to periodically run the function.   
4. The script automatically searches for Gmail threads matching the specified criteria and saves them as Google Docs in the specified folder. The title and URL of the Google document will be included in the spreadsheet specified by the properties; a message containing a link to the spreadsheet will be sent to the Google Space specified by the Webhook URL.  
