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

Sub LanguageIdentification()
'This subroutine sends text from cell A2 of Sheet1 to a language identification model, then returns the predicted language to cell B2 of Sheet1
Dim modelID As String
Dim versionID As String
Dim body As String
Dim res As String
Dim jobID As String
Dim outputJSON As String

'Model version and ID number for Modzy's "Language Identification" model which can identify 235 different languages
modelID = "bb63daa963"
versionID = "0.0.11"

'Hard coded example expects text in cell A2 of Sheet1 and will return a result in B2
body = "{" _
& "    ""model"":{" _
& "        ""identifier"": """ & modelID & """," _
& "        ""version"": """ & versionID & """" _
& "    }," _
& "    ""input"":{" _
& "        ""type"": ""text""," _
& "        ""sources"": {" _
& "            ""A2"": {" _
& "                 ""input.txt"": """ & CharCleanUp(Sheets(1).Range("A2").Value) & """" _
& "            }" _
& "        }" _
& "    }" _
& "}"

'This line sends an inference request to Modzy, using the above JSON as the input and captures the API response in the res variable
res = Modzy_API.ModzyJobSubmission(body)

jobID = Split(Split(res, ":")(5), ",")(0) 'This splits out the job id from JSON response
jobID = Left(jobID, Len(jobID) - 1) 'trimming off right quote mark
jobID = Right(jobID, Len(jobID) - 1) 'trimming off left quote mark for a final, clean jobID that we can use to get the results

'To-Do: Don't use a hard-coded wait time to check for results
If Application.Wait(Now + TimeValue("0:00:05")) Then
    outputJSON = Modzy_API.ModzyResults(jobID)
    Sheets(1).Range("B2") = Split(Split(outputJSON, "{")(4), "}")(0) 'This splits out the results.json portion of the results object
End If

End Sub

Function CharCleanUp(textToClean As String) As String
'This cleans up apostrophies that break the curl command
CharCleanUp = Replace(textToClean, "'", "'\''")
CharCleanUp = Replace(textToClean, "`", "'\`'")

End Function

Sub testcleanup()

Debug.Print CharCleanUp("I'm so happy")

End Sub
