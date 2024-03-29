Attribute VB_Name = "Modzy_API"
Option Explicit

Const modzyURL As String = "Modzy base URL goes here" ' e.g. "https://app.modzy.com"
Const APIKey As String = "your API key goes here" ' e.g. "u39fh3jf484hf89HFU9l.298vnLjwifjz08Lnwl82"

Private Declare PtrSafe Function popen Lib "libc.dylib" (ByVal command As String, ByVal mode As String) As LongPtr
Private Declare PtrSafe Function pclose Lib "libc.dylib" (ByVal file As LongPtr) As Long
Private Declare PtrSafe Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As Long
Private Declare PtrSafe Function feof Lib "libc.dylib" (ByVal file As LongPtr) As LongPtr

Function execShell(command As String, Optional ByRef exitCode As Long) As String
    ' execShell() function courtesy of Robert Knight via StackOverflow
    ' https://stackoverflow.com/questions/6136798/vba-shell-function-in-office-2011-for-mac

    Dim file As LongPtr
    file = popen(command, "r")

    If file = 0 Then
        Exit Function
    End If

    While feof(file) = 0
        Dim chunk As String
        Dim read As Long
        chunk = Space(50)
        read = fread(chunk, 1, Len(chunk) - 1, file)
        If read > 0 Then
            chunk = Left$(chunk, read)
            execShell = execShell & chunk
        End If
    Wend

    exitCode = pclose(file)
End Function

Function HTTPGet(sURI As String) As String

    Dim sCmd As String
    Dim sResult As String
    Dim lExitCode As Long

    sCmd = "curl -H ""Authorization: ApiKey " & APIKey & """ " & modzyURL & sURI
    sResult = execShell(sCmd, lExitCode)

    ' ToDo check lExitCode

    HTTPGet = sResult

End Function

Function HTTPPost(sURI As String, sData As String) As String

    Dim sCmd As String
    Dim sResult As String
    Dim lExitCode As Long

    sCmd = "curl -X POST -H ""Authorization: ApiKey " & APIKey & """ -H ""Content-Type: application/json"" -d '" & sData & "' " & modzyURL & sURI
    sResult = execShell(sCmd, lExitCode)

    ' ToDo check lExitCode

    HTTPPost = sResult

End Function

Function ModzyResults(jobID As String) As String
'This function GETs and returns a specific job result from Modzy
Dim route As String
Dim res As String

route = "/api/results/" & jobID
res = HTTPGet(route)
ModzyResults = res

End Function

Function ModzyJobSubmission(jobInput As String) As String
'This function POSTs a job request to Modzy and returns the API response from Modzy generated by this endpoint
Dim route As String
Dim res As String

route = "/api/jobs"
res = HTTPPost(route, jobInput)
ModzyJobSubmission = res

End Function

Function ModzyJobDetails(jobID As String) As String
'This function GETs a job’s details. It includes the status, total, completed, and failed number of items
Dim route As String
Dim res As String

route = "/api/jobs/" & jobID
res = HTTPGet(route)
ModzyJobDetails = res

End Function
