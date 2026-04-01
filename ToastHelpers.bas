Attribute VB_Name = "ToastHelpers"
' ToastHelpers.bas
' Version: 1.0
' Purpose: Simple, reliable toast notification functions
' Dependencies: Setup.bas, Logs.bas
' Date: 2025-10-27
Option Explicit


#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'=========================
' SIMPLE TOAST FUNCTIONS
'=========================

' Show a simple success toast
Public Sub ShowSuccessToast(ByVal Title As String, ByVal Message As String)
    SendToast Title, Message, "SUCCESS", 4, "BR", "?"
End Sub

' Show a simple info toast
Public Sub ShowInfoToast(ByVal Title As String, ByVal Message As String)
    SendToast Title, Message, "INFO", 4, "BR", "?"
End Sub

' Show a simple warning toast
Public Sub ShowWarningToast(ByVal Title As String, ByVal Message As String)
    SendToast Title, Message, "WARN", 5, "TR", "?"
End Sub

' Show a simple error toast
Public Sub ShowErrorToast(ByVal Title As String, ByVal Message As String)
    SendToast Title, Message, "ERROR", 6, "TR", "?"
End Sub

' Show a custom toast with full control
Public Sub ShowCustomToast( _
    ByVal Title As String, _
    ByVal Message As String, _
    Optional ByVal ToastType As String = "INFO", _
    Optional ByVal durationSec As Long = 4, _
    Optional ByVal Position As String = "BR", _
    Optional ByVal Icon As String = "")
    
    SendToast Title, Message, ToastType, durationSec, Position, Icon
End Sub

'=========================
' CORE TOAST SENDER
'=========================

Private Sub SendToast( _
    ByVal Title As String, _
    ByVal Message As String, _
    ByVal ToastType As String, _
    ByVal durationSec As Long, _
    ByVal Position As String, _
    ByVal Icon As String)
    
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim requestPath As String
    requestPath = fso.BuildPath(Setup.GetTempFolder(), "ToastRequest.json")
    
    ' Build JSON
    Dim json As String
    json = "{" & _
           """Title"": """ & EscapeJson(Title) & """," & _
           """Message"": """ & EscapeJson(Message) & """," & _
           """ToastType"": """ & UCase(ToastType) & """," & _
           """DurationSec"": " & durationSec & "," & _
           """Position"": """ & UCase(Position) & """"
    
    ' Add optional icon
    If Len(Icon) > 0 Then
        json = json & ",""Icon"": """ & Icon & """"
    End If
    
    json = json & "}"
    
    ' Write to file
    Dim ts As Object
    Set ts = fso.CreateTextFile(requestPath, True, False)
    ts.Write json
    ts.Close
    
    Logs.DebugLog "[ToastHelpers] Toast sent: " & Title, "INFO"
    Exit Sub
    
ErrorHandler:
    Logs.DebugLog "[ToastHelpers] Error sending toast: " & Err.Description, "ERROR"
    MsgBox "Failed to send toast: " & Err.Description, vbExclamation
End Sub

'=========================
' PROGRESS TOAST
'=========================

Public Sub ShowProgressToast( _
    ByVal Title As String, _
    ByVal Message As String, _
    ByVal ProgressPercent As Long, _
    Optional ByVal Position As String = "BR")
    
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim requestPath As String
    requestPath = fso.BuildPath(Setup.GetTempFolder(), "ToastRequest.json")
    
    ' Build JSON with progress
    Dim json As String
    json = "{" & _
           """Title"": """ & EscapeJson(Title) & """," & _
           """Message"": """ & EscapeJson(Message) & """," & _
           """ToastType"": ""INFO""," & _
           """DurationSec"": 0," & _
           """Position"": """ & UCase(Position) & """," & _
           """Progress"": " & ProgressPercent & _
           "}"
    
    ' Write to file
    Dim ts As Object
    Set ts = fso.CreateTextFile(requestPath, True, False)
    ts.Write json
    ts.Close
    
    Exit Sub
    
ErrorHandler:
    Logs.DebugLog "[ToastHelpers] Error sending progress toast: " & Err.Description, "ERROR"
End Sub

'=========================
' HELPER FUNCTIONS
'=========================

Private Function EscapeJson(ByVal s As String) As String
    ' Escape special characters for JSON
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJson = s
End Function

'=========================
' DEMO FUNCTIONS
'=========================

Public Sub DemoAllToasts()
    ShowInfoToast "Information", "This is an informational message"
    Sleep 2000
    
    ShowSuccessToast "Success!", "Operation completed successfully"
    Sleep 2000
    
    ShowWarningToast "Warning", "Please review this carefully"
    Sleep 2000
    
    ShowErrorToast "Error", "Something went wrong"
    Sleep 2000
    
    MsgBox "Demo complete!", vbInformation
End Sub

Public Sub DemoProgressToast()
    Dim i As Long
    For i = 0 To 100 Step 10
        ShowProgressToast "Processing Task", "Progress: " & i & "%", i
        Sleep 1000
    Next i
    
    ShowSuccessToast "Complete!", "Task finished successfully"
End Sub

