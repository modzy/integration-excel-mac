VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub SentimentAnalysis()
'This subroutine sends text from cell A2 of Sheet1 to a language identification model, then returns the predicted language to cell B2 of Sheet1
Dim modelID As String
Dim versionID As String
Dim outputFileName As String
Dim body As String
Dim res As String
Dim jobID As String
Dim outputJSON As String

'Model version and ID number for an open source Sentiment Analysis model available on Modzy's model library
modelID = "ed542963de"
versionID = "1.0.1"
outputFileName = "results.json" 'Each model has a specific name for its results file, check the API tab for the model you're working with to find it

'Hard coded example expects text in cell A2 of Sheet1 and will return a result in B2
body = "{" _
& "    ""model"":{" _
& "        ""identifier"": """ & modelID & """," _
& "        ""version"": """ & versionID & """" _
& "    }," _
& "    ""input"":{" _
& "        ""type"": ""text""," _
& "        ""sources"": {" _
& "            ""Cell A1"": {" _
& "                 ""input.txt"": """ & CharCleanUp(Sheets(1).Range("A1").value) & """" _
& "            }" _
& "        }" _
& "    }" _
& "}"

'This line sends an inference request to Modzy, using the above JSON as the input and captures the API response in the res variable
res = Modzy_API.ModzyJobSubmission(body)
jobID = GetJSONObjectValue(res, "jobIdentifier")

'Checks job status once a second and then downloads job results when status is "COMPLETED". Times out after 20 seconds
Dim time As Integer
time = 0
Do While time < 20
    If JobStatus(jobID) = "COMPLETED" Then
        outputJSON = Modzy_API.ModzyResults(jobID)
        Sheets(1).Range("B1") = SentimentAnalysisResultParsing(outputJSON, outputFileName)
        Exit Do
    End If
    Application.Wait (Now + TimeValue("0:00:01"))
    time = time + 1
Loop

'TO DO: Add in support to submit multiple inputs at a time

End Sub

Function SentimentAnalysisResultParsing(jobOutputJSON As String, outputFileName As String) As String
Dim results As String
Dim posScore As Double
Dim negScore As Double
Dim comboScore As Double

'A simple calculation which averages negative and positive scores to get a general sense of sentiment with a single number
results = Right(jobOutputJSON, Len(jobOutputJSON) - InStr(jobOutputJSON, outputFileName) + 2)
negScore = Split(Split(results, ":")(10), "}")(0)
posScore = Split(Split(results, ":")(8), "}")(0)
comboScore = (posScore + negScore * (-1)) / 2

SentimentAnalysisResultParsing = comboScore

End Function

Function JobStatus(jobID As String) As String
Dim status As String
'Get the job status
status = Modzy_API.ModzyJobDetails(jobID)
'Look for the "status" object and trim out everything before it
status = GetJSONObjectValue(status, "status")
JobStatus = status

End Function

Function GetJSONObjectValue(json As String, objectName As String)
'This function returns the value associated with any objectname found within the provided json (i.e. given '"key": "value"', search for "key", return "value")
Dim value As String
value = Right(json, Len(json) - InStr(json, objectName) + 2) 'Look for the object name provided and trim out everything before it
value = Split(Split(value, ":")(1), ",")(0) 'This splits out the object value from the rest of the JSON
value = Left(value, Len(value) - 1) 'trimming off right quote mark
value = Right(value, Len(value) - 1) 'trimming off left quote mark for a final, clean value
GetJSONObjectValue = value

End Function

Function CharCleanUp(textToClean As String) As String
'This cleans up apostrophies that break the curl command
CharCleanUp = Replace(textToClean, "'", "'\''")
CharCleanUp = Replace(textToClean, "`", "'\`'")

End Function
