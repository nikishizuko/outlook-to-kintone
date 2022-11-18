# Outlook to Kintone Add-in

## Overview: 
This add-in for Outlook allows the user to send email data to Kintone as a new record.  
Note that this set up requires a Make scenario and Webhook. 

## Setup: 
### Kintone app set up
**Required fields:**
  - Subject (Text)
  - To recipients (Text)
  - CC recipients (Text)
  - BCC recipients (Text)
  - Email body (Text Area)

Note that field codes and field names do not matter because mapping will be done in Make.

### Make set up
1. Trigger is "Custom Webhook" module. Save the webhook URL.
2. Second step is "Parse JSON" module
  - JSON string should be **Value**
3. Final step is "Create a record" Kintone module
  - App ID: App ID of Kintone app
  - Subject value: **Subject** of Parse JSON module
  - BCC Recipients value: **recipients.bcc[]** of Parse JSON module
  - To Recipients value: **recipients.to[]** of Parse JSON module
  - CC Recipients value: **recipients.cc[]** of Parse JSON module
  - Email Body value: **bodyText** of Parse JSON module

### Add-in set up
1. Save add-in folder contents to local computer
2. In Outlook client or browser version, go to "Get Add-ins"
3. Go to "My add-ins"
4. Under "Custom addins" click "Add a custom add-in" and "Add from File"
5. Select the **manifest.xml** file from the source folder
6. With the add-in installed, open the add-in by going to the message compose screen and clicking "Outlook email to Kintone" from the add-in list
7. Click the gear icon, paste in the Webhook URL from Make 

## Usage
1. In the message compose screen, copy email body text and paste into the taskpane text area.
2. Click "Send" to send data and create a new record with Subject, Recipients, and Body Text data in Kintone.
