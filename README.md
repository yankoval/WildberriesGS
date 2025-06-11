# Google Apps Script for Sheet Automation and Monitoring

This project utilizes Google Apps Script to automate various tasks within Google Sheets. Its main capabilities include copying data between sheets, archiving data, and monitoring data with notifications.

## Core Functionalities

### Data Copying and Archiving (`copyDataFromAnotherSheet()`)
This function copies data from a specified source Google Sheet to a target Google Sheet.
- **Source Spreadsheet ID:** `16RKOu7ALQY805E1FX-aX9a1sLV27J1bQf1r431FIk8E`
- **Source Sheet Name:** 'Короба'
- **Target Spreadsheet ID:** `1vaDAM2qoeBTlsBOI5Z3NYFpDIK8QlnF1muCYU8qeH6o`
- **Target Sheet Name:** 'Лист1'

Additionally, it archives the copied data by creating a new sheet in the target spreadsheet. The new sheet is named 'Архив_YYYY-MM-DD_HH-mm-ss', where YYYY-MM-DD is the current date and HH-mm-ss is the current time.

### Data Monitoring and Notification (`checkForPositiveValues()`)
This function checks a specific column for positive numerical values and triggers notifications if certain conditions are met.
- It monitors column 2 (B) in the 'Короба' sheet (ID: `16RKOu7ALQY805E1FX-aX9a1sLV27J1bQf1r431FIk8E`).
- If a positive value is found in column 2, it retrieves the corresponding text from column 1 (A) of the same row.
- This retrieved text is then compared against a list of values in a named range 'СписокРЦ', which is located in a different spreadsheet (ID: `1vaDAM2qoeBTlsBOI5Z3NYFpDIK8QlnF1muCYU8qeH6o`).

If a match is found between the retrieved text and a value in 'СписокРЦ', the following notifications are triggered:
- An email is sent to 'yankoval@gmail.com' with the matched value(s) in the subject line.
- A toast message is displayed in the active spreadsheet.
- A browser message box is displayed.

Error handling is implemented to catch and display any errors in a message box and log them for troubleshooting.

## Helper Functions

### `createArchive(data)`
This function is utilized by `copyDataFromAnotherSheet()` to archive data. It creates a new sheet within the active spreadsheet. The sheet is given a timestamped name in the format 'Архив_YYYY-MM-DD_HH-mm-ss' (e.g., 'Архив_2023-10-27_10-30-00'). The `data` provided to the function (which is the data copied from the source sheet) is then populated into this newly created archive sheet.

### `sendSimpleEmail(txtRC)`
This function is responsible for dispatching email notifications.
- The recipient's email address is currently hardcoded as 'yankoval@gmail.com'.
- The subject line of the email dynamically includes the `txtRC` parameter. In the context of `checkForPositiveValues()`, `txtRC` represents the value(s) that matched the criteria, thereby alerting the recipient to the specific data that triggered the notification.

### `deleteExistingTriggers()`
This function serves to remove all existing script triggers associated with the current Google Apps Script project. Its primary use case is within the `createTimeDrivenTrigger()` function to ensure that duplicate or conflicting time-based triggers are not created, maintaining a clean and predictable trigger setup.

## Automated Execution (Triggers)

### `createTimeDrivenTrigger()`
This function is responsible for setting up automated execution of a specific function at regular intervals.
- **Action:** It first calls `deleteExistingTriggers()` to remove any previously set triggers for the project. This prevents the accumulation of duplicate triggers.
- After clearing existing triggers, it creates a new time-driven trigger.
- **Scheduled Function:** The trigger is configured to run a function named `'checkWildberries'`.
- **Frequency:** This execution is scheduled to occur every 10 minutes.

**Note:** The function `'checkWildberries'` is referenced by this trigger but is not explicitly defined in the provided `Код.gs` script. It is presumed that `'checkWildberries'` might be an alias for, or is intended to be, the `checkForPositiveValues()` function, which handles the data monitoring and notification logic.

## Configuration and Setup

To adapt this script to your specific Google Sheets environment, you will need to modify several hardcoded values within the script (`Код.gs`). Open the script editor in your Google Sheet to make these changes.

Below is a list of key hardcoded variables and where to find them:

**1. In `copyDataFromAnotherSheet()` function:**
   - `sourceSpreadsheetId`: The ID of the Google Spreadsheet from which data will be copied.
     *(Currently: `'16RKOu7ALQY805E1FX-aX9a1sLV27J1bQf1r431FIk8E'`)*
   - `sourceSheetName`: The name of the specific sheet within the source spreadsheet (e.g., `'Короба'`).
     *(Currently: `'Короба'`)*
   - `targetSpreadsheetId`: The ID of the Google Spreadsheet where data will be pasted and archived.
     *(Currently: `'1vaDAM2qoeBTlsBOI5Z3NYFpDIK8QlnF1muCYU8qeH6o'`)*
   - `targetSheetName`: The name of the sheet in the target spreadsheet where data will be pasted (e.g., `'Лист1'`).
     *(Currently: `'Лист1'`)*

**2. In `checkForPositiveValues()` function:**
   - `spreadsheetId`: The ID of the Google Spreadsheet that this function will primarily monitor.
     *(Currently: `'16RKOu7ALQY805E1FX-aX9a1sLV27J1bQf1r431FIk8E'`)*
   - `listSpreadsheetId`: The ID of the Google Spreadsheet that contains the named range `'СписокРЦ'` for cross-referencing.
     *(Currently: `'1vaDAM2qoeBTlsBOI5Z3NYFpDIK8QlnF1muCYU8qeH6o'`)*
   - `sheetName`: The name of the sheet to be checked for positive values (e.g., `"Короба"`).
     *(Currently: `"Короба"`)*
   - `columnToCheck`: The column number (e.g., B=2) in `sheetName` that will be scanned for positive numerical values.
     *(Currently: `2`)*
   - `targetColumn`: The column number (e.g., A=1) in `sheetName` from which corresponding text is retrieved if a positive value is found in `columnToCheck`.
     *(Currently: `1`)*
   - `listRangeName`: The name of the range (e.g., `"СписокРЦ"`) in `listSpreadsheetId` that contains values to check against.
     *(Currently: `"СписокРЦ"`)*

**3. In `sendSimpleEmail()` function:**
   - `recipient`: The email address to which notifications will be sent.
     *(Currently: `'yankoval@gmail.com'`)*

**Instructions for Users:**
- Carefully review each of these variables in the script.
- Replace the existing hardcoded values with the IDs, sheet names, column numbers, and email addresses relevant to your Google Sheets and desired workflow.
- Ensure that the named ranges (like `'СписокРЦ'`) exist in your specified spreadsheets.
- After making changes, save the script in the Google Apps Script editor.
- You may also need to grant necessary permissions when the script runs for the first time after modifications.

## How to Use

Follow these steps to get the script up and running:

**1. Accessing the Script Editor:**
   - This script is written in Google Apps Script, which is integrated with Google Sheets.
   - Open the Google Sheet to which you want to attach this script (or the one that already contains it).
   - Navigate to **Tools > Script editor**. This will open the Google Apps Script editor in a new tab or window, where you can paste or view the `Код.gs` script.

**2. Manual Execution of Functions:**
   - You can run individual functions directly from the script editor for testing or one-time operations.
   - In the script editor, there's a dropdown menu near the top (often labeled "Select function").
   - Choose the function you wish to run (e.g., `copyDataFromAnotherSheet()` or `checkForPositiveValues()`).
   - Click the "Run" button (it looks like a play icon ▶).
   - The script will execute, and any UI messages (like toasts or message boxes) will appear in the context of the Google Sheet from which you opened the script editor.

**3. Setting Up Automated Execution:**
   - To have the script run automatically (e.g., to periodically check for values), you need to set up a time-driven trigger.
   - The function `createTimeDrivenTrigger()` is designed for this.
   - Manually run `createTimeDrivenTrigger()` once from the script editor (as described in step 2). This will:
     - Delete any existing triggers for the project by calling `deleteExistingTriggers()`.
     - Create a new trigger that runs the function `'checkWildberries'` every 10 minutes.
   - **Important Note:** As mentioned in the "Automated Execution (Triggers)" section, `'checkWildberries'` is the function name set by the trigger. Ensure this function exists or is an alias for `checkForPositiveValues()` as intended.

**4. Granting Permissions:**
   - The first time you run a function that requires access to your Google services (like reading/writing to Sheets or sending emails), Google will prompt you for authorization.
   - A dialog box will appear asking you to "Review Permissions."
   - Follow the prompts to select your Google account and grant the necessary permissions.
   - This is a standard security measure for Google Apps Script to ensure you authorize the script to act on your behalf. You typically only need to do this once, unless you modify the script to require new types of permissions.

After setup and authorization, the script will perform its tasks according to the functions executed, either manually or via the automated trigger. Remember to configure the hardcoded values as detailed in the "Configuration and Setup" section to tailor the script to your needs.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.
