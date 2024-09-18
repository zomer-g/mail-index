# Email Indexing Script

## Overview
This Google Apps Script automatically indexes new emails from your Gmail inbox, organizes attachments into dedicated folders in Google Drive, and logs email details into a Google Sheet. The script prompts for necessary IDs when run, making it easy to customize for different users.

## Features
- **Index New Emails**: Processes new emails based on the last indexed date.
- **Organize Attachments**: Creates subfolders in Google Drive for email attachments.
- **Detailed Logging**: Records email details, including date, sender, clear email address, subject, and links.
- **Chronological Processing**: Indexes emails starting with the oldest.

## Requirements
- A Google Account with access to Gmail, Google Drive, and Google Sheets.
- Authorization for the script to access these services.

## Setup Instructions

1. **Open Google Sheets**:
   - Create a new Google Sheet to log email details.

2. **Open Apps Script**:
   - In your Google Sheet, click on **Extensions** > **Apps Script**.

3. **Copy the Script**:
   - Paste the provided script into the Apps Script editor.

4. **Save and Authorize**:
   - Save the script and run it for the first time to authorize access to your Gmail, Google Drive, and Google Sheets.

5. **Run the Script**:
   - Enter the Google Sheet ID and Google Drive folder ID when prompted.

## Configuring a Trigger
To automate the script, set up a time-driven trigger:

1. In the Apps Script editor, click on the clock icon (Triggers).
2. Click on **"+ Add Trigger"**.
3. Select `indexNewEmails` from the function dropdown.
4. For "Select event source," choose **Time-driven**.
5. Choose the frequency for the trigger (e.g., every hour).
6. Click **Save**.

## Usage
- Run the script from the Apps Script editor or allow the trigger to execute it automatically.
- The script will log new emails since the last indexing, along with their details and any attachments.

## Notes
- Ensure that the specified Google Sheet and Google Drive folder are accessible.
- The script indexes only emails that are newer than the last recorded email in the Google Sheet.

## Troubleshooting
- Check permissions granted to the script if you encounter issues.
- Review logs for error messages related to email or Drive access.

## License
This script is provided for free use. Modify it as needed for personal or educational purposes.
