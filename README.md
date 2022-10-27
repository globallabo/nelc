# NELC Curriculum Generator

Using Google Apps Script, it is possible to automatically generate documents based on a template made with Google Docs and content stored in Google Sheets.

The script should be created on the Google Sheets spreadsheet. The Google Docs templates and Drive folders should be supplied as environment variables by using [the Properties Service](https://developers.google.com/apps-script/guides/properties#manage_script_properties_manually).

First, this script will create a new menu in the Google Sheets UI. The menu will contain an option to run the automation. When that is selected, the rest of the script runs.

This will create a new Google Docs document for each row in the spreadsheet, which represents one lesson in the curriculum. In our case, there are two levels, which are represented by two tabs in the spreadsheet.

Additionally, Google Drive is used to save the Google Docs documents as PDF files directly into a Drive folder. When changes have been made, the output files are updated by first removing the previous copy and then saving a new copy. Simply moving to a user's trash folder is the obvious approach, but in order to allow for multiple users of the script, it's better to supply a shared "trash" folder located elsewhere in Drive.
