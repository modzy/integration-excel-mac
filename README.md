# integration-excel-mac
Resources for building Modzy integrations into Excel for Mac with VBA

## How to use this integration with the example code provided
1. *Download files*: Clone this repo, or download Modzy_API.bas and LanguageDetectionExample.cls
2. *Enable your Developer Tab*: Open up Excel -> go to Preferences -> go to Ribbon & Toolbar -> click on the "Developer" tab to enable it in the main ribbon
3. *Open up the VBA IDE*: Click on the "Developer" tab -> Click on the "Visual Basic" icon
4. *Import API Module*: In the "Project - VBA Project" window right click on the "Modules" folder -> click on "Import File" -> select Modzy_API.bas
5. *Import Seniment Analysis Example*: In the "Project - VBA Project" window right click on "Sheet1" -> click on "Import File" -> select SenimentAnalysisExample.cls
6. *Update environment variables*: At the top of the Modzy_API module, update the URL of your instance of Modzy, along with the API Key you'll be using to call Modzy
7. *Add sample input*: Add any text you'd like to Cell "A1" on Sheet1 of your spreadsheet
8. *Run a sample inference*: In the VBA IDE, double click on the "Sheet1" object -> click your mouse somewhere within the "Sub LanguageIdentification()" subroutine -> Click on the triangular run button at the top of the editor

## Modzy_API Functions
The Modzy_API module lets you interact with the Modzy API using CURL. A lot of the annoyance of setting up, executing, and returning a result has been abstracted away. The main functions you'll likely want to use are:

`ModzyJobSubmission`
This function accepts a JSON string of the full request you'd like to submit to Modzy and returns the reponse provided by the /api/jobs endpoint. This response includes the JobID generated when a job is successfully submitted to Modzy which can be used to query the result of that job.

`ModzyResults`
This function accepts a valid Job ID and returns the results of any inference that has been successfully completed.

`ModzyJobDetails`
This function GETs a job’s details. It includes the status, total, completed, and failed number of items
