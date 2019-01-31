# Integration of jira with google spreadsheets
Project that allows you to fetch the data from JIRA API and update Google Sheets spreadsheet with the data about the last completed sprint.

## Getting Started

### Prerequisites
To programmatically access your spreadsheet, you’ll need to create a service account and OAuth2 credentials from the [Google API Console](https://console.developers.google.com):
* Create a new project
* Click Enable API. Search for and enable the Google Drive API
* Create credentials for a Web Server to access Application Data
* Name the service account and grant it a Project Role of Editor
* Download the JSON file
* Copy the JSON file to directory session inside of your project and rename file to client_secret.json

Use pip (package manager) to install all requirements from requirements.txt

```
pip install -r requirements.txt
```

Add environment variables to your project. You can do this using ide or features of your operating system.
## Running

Run run.py file

For windows:
```
python run.py
```
For unix-like operating systems:
```
python3 run.py
```

