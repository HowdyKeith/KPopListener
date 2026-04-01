Attribute VB_Name = "MsgBoxToastAPI"
'***************************************************************
' Module: MsgBoxToastAPI
' Version: 7.1
' Purpose: Display Excel/Office Toasts via Native Windows Toast API
' Dependencies: clsToastNotification.cls (v11.8), Logs.bas (v1.0.1), MsgBoxUnified.bas (v1.0.2), MsgBoxUniversal.bas (v3.3)
' Features:
'   - Queued & stacked toasts
'   - Progress updates
'   - Native sound alerts
'   - Callback macro auto-invoker
'   - ImagePath support
' Changes:
'   - Fixed incorrect method call from currentToast.Close to currentToast.CloseToast
'   - Added ImagePath parameter to ShowToast
'   - Updated to use clsToastNotification v11.8
'   - Added logging via Logs.DebugLog
' Updated: 2025-10-24
'***************************************************************
Option Explicit

Private SharedQueue As Collection

Private Sub InitQueue()
    If SharedQueue Is Nothing Then Set SharedQueue = New Collection
End Sub

Public Sub ShowToast(ByVal Title As String, ByVal Message As String, _
                     Optional ByVal Level As String = "INFO", _
                     Optional ByVal Duration As Long = 5, _
                     Optional ByVal Position As String = "BR", _
                     Optional ByVal SoundName As String = "", _
                     Optional ByVal CallbackMacro As String = "", _
                     Optional ByVal ImagePath As String = "")
    On Error GoTo ErrorHandler
    Dim Toast As clsToastNotification
    Set Toast = New clsToastNotification
    
    Toast.Title = Title
    Toast.Message = Message
    Toast.Level = Level
    Toast.Duration = Duration
    Toast.Position = Position
    Toast.SoundName = SoundName
    Toast.CallbackMacro = CallbackMacro
    Toast.ImagePath = ImagePath
    
    InitQueue
    SharedQueue.Add Toast
    Logs.DebugLog "[MsgBoxToastAPI] Queued toast: " & Title, "INFO"
    If SharedQueue.count = 1 Then DisplayNextToast
    
    Exit Sub
ErrorHandler:
    Logs.DebugLog "[MsgBoxToastAPI] Error in ShowToast: " & Err.Description, "ERROR"
End Sub

Private Sub DisplayNextToast()
    On Error GoTo ErrorHandler
    If SharedQueue.count = 0 Then Exit Sub
    
    Dim currentToast As clsToastNotification
    Set currentToast = SharedQueue(1)
    
    ' Use PowerShell delivery (as per original intent)
    currentToast.Show "powershell"
    
    ' Wait for duration
    Dim waitSeconds As Long
    waitSeconds = currentToast.Duration
    If currentToast.IsShowing And waitSeconds > 0 Then
        Application.Wait Now + TimeSerial(0, 0, waitSeconds)
    End If
    
    ' Close the toast
    currentToast.CloseToast
    SharedQueue.Remove 1
    Logs.DebugLog "[MsgBoxToastAPI] Displayed and closed toast: " & currentToast.Title, "INFO"
    
    If SharedQueue.count > 0 Then DisplayNextToast
    
    Exit Sub
ErrorHandler:
    Logs.DebugLog "[MsgBoxToastAPI] Error in DisplayNextToast: " & Err.Description, "ERROR"
    If SharedQueue.count > 0 Then
        SharedQueue.Remove 1
        If SharedQueue.count > 0 Then DisplayNextToast
    End If
End Sub

'***************************************************************
' Sample Usage
'***************************************************************
Public Sub TestToastAPI()
    On Error Resume Next
    ShowToast "Test Toast", "This is a test notification.", "INFO", 5, "BR", "", "", "C:\Images\icon.png"
    ShowToast "Queued Toast", "This should appear next.", "SUCCESS", 3, "TR", "", ""
    ShowToast "Error Toast", "Something went wrong!", "ERROR", 5, "C", "", ""
End Sub




