# 🧹 Public Computer Cleaner

Cleans public Windows computers (such as those in libraries or hotels) of personal information. Can be scheduled to run automatically.

## Features

* Deletes all files and folders from Desktop, Documents, Pictures, Videos, Music, and Downloads.
* Kills running instances of Chrome and Edge.
* Cleares all Chrome and Edge user data.
* Tracks the number of files and folders deleted in each case.
* Computers are able to be managed from a Google Sheet.
* Sends a monthly report containing information about the files/folders deleted from each computer at the center.
* Logs to an output file.
* Able to be scheduled to run automatically.
* Sets the background of the computer to a chosen image.

## Installation

1. Download this repository locally onto all of the public computers that are going to be cleaned/managed.
2. Create a copy of the Public Computer Cleaner Google Sheet that is used to manage the computers. Do that [here](https://docs.google.com/spreadsheets/d/1ImUUp0SANELihf5KzvDEwFottb22u_uVLanmSsI9-N4/copy?usp=sharing).
3. Make sure a version of Python 3 is installed. The program was developed using Python 3.8.5, so there could be possible issues with using earlier editions of Python 3.
4. For a better experience configuring these files on the public computers, consider installing an IDE, such as Visual Studio Code.

## Configuration Step 1: Google Sheet

The steps for configuring the Google Sheet can be found in the Google Sheet itself (specifically in the *Controls* sheet).

## Configuration Step 2: Configuring `gspread`

[`gspread`](https://github.com/burnash/gspread) is the module used to communicate with the Public Computer Cleaner Google Sheet from the program. This module needs proper authentication to be able to successfully communicate with the sheet.

To authenticate for this version of gspread:
1. Head over to the [Google Cloud Developers Console](https://console.developers.google.com/)
2. Create/select a project.
3. In the API library, enable the Google Drive API and the Google Sheets API.
4. Head to APIs & Services > Credentials and click the "Create Credentials" button. Click "Service Account".
5. Fill out the form and save.
6. Revisit the service account you just created and click "Create new key" under "Keys".
7. Select JSON and save this file as `service_account.json`. Move the file to the Public Computer Cleaner directory.
8. Share the Google Sheet you previously copied with the email in the `client-email` field of the JSON file. This allows `gspread` to access the Google Sheet.

If you need help, you might find this [gspread help article](https://gspread.readthedocs.io/en/latest/oauth2.html) helpful. Remember that this program runs on version 3.6.0 of `gspread`, so the documentation may or may not be ahead.

## Configuration Step 3: Local Files

1. In `windows.py`, edit `imgStr`, the path of the image you wish to have set as the background of the computer on run.
2. In `cleaner.py`, change the first 6 variables in the file (`documentsStr`, `downloadsStr`, etc) to match the paths of the relevant locations in the computer. For example:

    ```python
    documentsStr = r"C:\Users\jdoe\Documents"
    desktopStr = r"C:\Users\jdoe\Desktop"
    ```

    The reason these are customizable is to provide a test environment for developing or managing. For example, you could customize the paths to a custom folder, put files in the custom folder, and run the program to check if it is deleting the appropriate files.
3. In `chrome.py`, change the `chromeDataStr` variable to be the path to the Chrome user data folder.

    To find this, open Google Chrome and visit `chrome://version`. Look for the `Profile Path` field. The user data folder is the parent of that directory.

    For example:

    ```python
    chromeDataStr = r'C:\Users\jdoe\AppData\Local\Google\Chrome\User Data'
    ```

4. Make sure that the version of Edge on the computer is v79 or greater. Then, in `edge.py`, fill your computer's username in `edgeDataStr` (such as `jdoe`).

5. If you have not already, create `settings.txt` and add

    ```
    COMPUTER_NAME=Computer 1
    ```

    `Computer 1` should be replaced with the name of the computer as you inputted into your Google Sheet. You are welcome to add other local settings if you like.

6. In `sheets.py`, edit `spreadsheetName` to be the name of your Google Sheet in Google Drive.

7. In `emailer.py` edit `sheetUrl` to be the URL of the Google Sheet you copied.

8. In a terminal in the same directory, install the pip requirements. [Help on that.](https://pip.pypa.io/en/latest/user_guide/#requirements-files)

9. Test run `manager.py`, and note if anything is not set correctly. If everything works, you are welcome to schedule running `manager.py`.

## Contributing

Contributions are welcome.