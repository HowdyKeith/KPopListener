Attribute VB_Name = "MsgBoxUniversal"
'***************************************************************
' Module: MsgBoxUniversal
' Version: 3.13
' Purpose: Universal Toast Manager for VBA
' Author: Keith Swerling + Claude
' Dependencies: Logs.bas (v1.0.3), MsgBoxMSHTA.bas (v6.15), Setup.bas (v1.8),
'               ToastSender.bas (v5.5), clsToastNotification.cls (v11.20),
'               MsgBoxMain.bas (v11.10), MsgBoxPython.bas (v1.9)
' Features:
'   - Unified toast delivery (ShowToast) with ToastSender.bas integration
'   - Named pipe priority with automatic temp file fallback
'   - Quick toast methods (MsgInfoEx, MsgWarnEx, MsgErrorEx, MsgSuccessEx)
'   - JSON and HTML escaping utilities
'   - Simplified listener status checks for KPopListener
' Changes:
'   - v3.13: Updated for KPopListener naming (pipe, listeners, sentinel files)
'   - v3.12: Integrated ToastSender.bas for pipe/temp communication
'   - v3.12: Simplified PowershellListenerRunning using native pipe check
'   - v3.11: Enhanced WMI query, increased timeout
' Updated: 2025-10-29
'***************************************************************
Option Explicit

' Sleep helper
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


#If VBA7 Then
    Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
        ByVal lpFileName As String, _
        ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, _
        ByVal lpSecurityAttributes As Long, _
        ByVal dwCreationDisposition As Long, _
        ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) As Long
    
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
#Else
    Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
        ByVal lpFileName As String, _
        ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, _
        ByVal lpSecurityAttributes As Long, _
        ByVal dwCreationDisposition As Long, _
        ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) As Long
    
    Private Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
#End If

' Constants for CreateFile
Private Const GENERIC_WRITE As Long = &H40000000
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const PIPE_PATH As String = "\\.\pipe\ExcelToastPipe"

'=========================
' CONFIGURATION
'=========================
Private Const REQUEST_FILE As String = "ToastRequest.json"
Private Const STATUS_FILE As String = "ToastListenerStatus.json"

'=========================
' LISTENER CONTROL
'=========================
Public Sub StartToastListener()
    On Error Resume Next
    
    If Not Setup.AutoStartToastServers Then
        Logs.LogInfo "[MsgBoxUniversal] Autostart disabled; manually start KPopListener.ps1"
        Exit Sub
    End If
    
    ' Use ToastSender's built-in starter
    ToastSender.StartToastListener
    Logs.LogInfo "[MsgBoxUniversal] StartToastListener delegated to ToastSender"
End Sub

Public Sub StopToastListener()
    On Error GoTo ErrHandler
    Dim wmi As Object: Set wmi = GetObject("winmgmts://./root/cimv2")
    Dim query As String
    Dim processes As Object
    Dim proc As Object
    Dim terminated As Long
   
    ' Stop PowerShell listener (KPopListener)
    query = "SELECT * FROM Win32_Process WHERE Name = 'powershell.exe' AND CommandLine LIKE '%KPopListener%.ps1%'"
    Set processes = wmi.ExecQuery(query)
    terminated = 0
    For Each proc In processes
        proc.Terminate
        terminated = terminated + 1
    Next
    Logs.LogInfo "[MsgBoxUniversal] Terminated " & terminated & " PowerShell listener(s)"
   
    ' Stop Python listener (KPopListener)
    query = "SELECT * FROM Win32_Process WHERE Name = 'python.exe' AND CommandLine LIKE '%KPopListener%.py%'"
    Set processes = wmi.ExecQuery(query)
    terminated = 0
    For Each proc In processes
        proc.Terminate
        terminated = terminated + 1
    Next
    Logs.LogInfo "[MsgBoxUniversal] Terminated " & terminated & " Python listener(s)"
   
    MsgBox "KPopListener stopped." & vbCrLf & vbCrLf & _
           "PowerShell: " & terminated & " process(es)" & vbCrLf & _
           "Python: " & terminated & " process(es)", vbInformation, "Listeners Stopped"
    Exit Sub
    
ErrHandler:
    Logs.LogError "[MsgBoxUniversal] StopToastListener error: " & Err.Description
    MsgBox "Error stopping listeners: " & Err.Description, vbCritical
End Sub

Public Function PowershellListenerRunning() As Boolean
    On Error Resume Next
    
    ' Fast check: Try to connect to named pipe
    Dim hPipe As Long
    hPipe = CreateFile(PIPE_PATH, GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    
    If hPipe <> INVALID_HANDLE_VALUE Then
        CloseHandle hPipe
        PowershellListenerRunning = True
        Logs.LogDebug "[MsgBoxUniversal] KPopListener (PowerShell) running (pipe available)"
    Else
        PowershellListenerRunning = False
        Logs.LogDebug "[MsgBoxUniversal] KPopListener (PowerShell) not running (pipe unavailable)"
    End If
End Function

Public Function PythonListenerRunning() As Boolean
    On Error GoTo ErrHandler
    Dim wmi As Object
    Set wmi = GetObject("winmgmts://./root/cimv2")
    Dim query As String
    query = "SELECT * FROM Win32_Process WHERE Name = 'python.exe' AND CommandLine LIKE '%KPopListener%.py%'"
    Dim processes As Object
    Set processes = wmi.ExecQuery(query)
    Dim processCount As Long: processCount = processes.count
    Logs.LogDebug "[MsgBoxUniversal] KPopListener (Python) check: Found " & processCount & " matching processes"
    PythonListenerRunning = processCount > 0
    Exit Function

ErrHandler:
    Logs.LogError "[MsgBoxUniversal] PythonListenerRunning error: " & Err.Description
    PythonListenerRunning = False
End Function

'=========================
' TOAST DELIVERY (Integrated with ToastSender.bas)
'=========================
Public Function ShowToast( _
    ByVal Title As String, _
    ByVal Message As String, _
    ByVal Level As String, _
    ByVal durationSec As Long, _
    ByVal Position As String, _
    ByVal Sound As String, _
    ByVal ImagePath As String, _
    ByVal IsProgress As Boolean, _
    ByVal ProgressValue As Long) As Boolean
    
    On Error GoTo ErrHandler
    
    ' Check if toast system is ready
    If Not MsgBoxMain.IsToastSystemReady Then
        Logs.LogWarn "[MsgBoxUniversal] Toast system not ready; falling back to MSHTA"
        MsgBoxMSHTA.ShowToast Title, Message, Level, durationSec, Position, Sound, ImagePath
        ShowToast = True
        Exit Function
    End If
    
    ' Try PowerShell listener via ToastSender (pipe first, then temp file)
    If PowershellListenerRunning() Then
        ' ToastSender.SendToast will try named pipe first, auto-fallback to temp file
        ToastSender.SendToast Title, Message, Level, durationSec, , , "Pipe"
        ShowToast = True
        Logs.LogInfo "[MsgBoxUniversal] Toast sent via ToastSender (Pipe?Temp): " & Title
        Exit Function
    End If
    
    ' Try Python listener if available
    If PythonListenerRunning() Then
        On Error Resume Next
        MsgBoxPython.SendPythonNotification Title, Message, "", durationSec, Position, ImagePath, Level, ProgressValue
        If Err.Number = 0 Then
            ShowToast = True
            Logs.LogInfo "[MsgBoxUniversal] Toast sent via Python: " & Title
            Exit Function
        End If
        Logs.LogWarn "[MsgBoxUniversal] Python toast failed: " & Err.Description
        On Error GoTo ErrHandler
    End If
    
    ' Final fallback: MSHTA
    MsgBoxMSHTA.ShowToast Title, Message, Level, durationSec, Position, Sound, ImagePath
    ShowToast = True
    Logs.LogInfo "[MsgBoxUniversal] MSHTA fallback toast displayed: " & Title
    Exit Function

ErrHandler:
    Logs.LogError "[MsgBoxUniversal] ShowToast error: " & Err.Description
    ShowToast = False
End Function

'=========================
' QUICK TOAST METHODS
'=========================
Public Sub MsgInfoEx(ByVal Message As String, Optional ByVal Position As String = "BR")
    Dim t As New clsToastNotification
    With t
        .Title = "Info"
        .Message = Message
        .Level = "INFO"
        .Duration = 5
        .Position = Position
        .Icon = "?"
        .YOffset = 0
    End With
    t.Show "auto"
    Logs.LogInfo "[MsgBoxUniversal] MsgInfoEx: " & Message
End Sub

Public Sub MsgWarnEx(ByVal Message As String, Optional ByVal Position As String = "TR")
    Dim t As New clsToastNotification
    With t
        .Title = "Warning"
        .Message = Message
        .Level = "WARN"
        .Duration = 6
        .Position = Position
        .Icon = "?"
        .SoundName = "BEEP"
        .YOffset = 0
    End With
    t.Show "auto"
    Logs.LogInfo "[MsgBoxUniversal] MsgWarnEx: " & Message
End Sub

Public Sub MsgErrorEx(ByVal Message As String, Optional ByVal Position As String = "TL")
    Dim t As New clsToastNotification
    With t
        .Title = "Error"
        .Message = Message
        .Level = "ERROR"
        .Duration = 8
        .Position = Position
        .Icon = "?"
        .SoundName = "BEEP"
        .YOffset = 0
    End With
    t.Show "auto"
    Logs.LogInfo "[MsgBoxUniversal] MsgErrorEx: " & Message
End Sub

Public Sub MsgSuccessEx(ByVal Message As String, Optional ByVal Position As String = "BR")
    Dim t As New clsToastNotification
    With t
        .Title = "Success"
        .Message = Message
        .Level = "SUCCESS"
        .Duration = 4
        .Position = Position
        .Icon = "?"
        .YOffset = 0
    End With
    t.Show "auto"
    Logs.LogInfo "[MsgBoxUniversal] MsgSuccessEx: " & Message
End Sub

'=========================
' PROGRESS TOAST (Using ToastSender)
'=========================
Public Function ShowToastWithProgress( _
    ByVal Title As String, _
    ByVal Message As String, _
    ByVal ProgressPercent As Long, _
    Optional ByVal ToastType As String = "INFO", _
    Optional ByVal Duration As Long = 5) As Boolean
    
    On Error GoTo ErrHandler
    
    If PowershellListenerRunning() Then
        ' Use ToastSender for progress toasts
        ' Note: KPopListener.ps1 supports progress in the JSON
        ToastSender.SendToast Title, Message & " (" & ProgressPercent & "%)", ToastType, Duration, , , "Pipe"
        ShowToastWithProgress = True
        Logs.LogInfo "[MsgBoxUniversal] Progress toast sent: " & Title & " - " & ProgressPercent & "%"
    Else
        ' Fallback to regular toast
        ShowToast Title, Message & " (" & ProgressPercent & "%)", ToastType, Duration, "BR", "", "", False, ProgressPercent
        ShowToastWithProgress = True
    End If
    Exit Function

ErrHandler:
    Logs.LogError "[MsgBoxUniversal] ShowToastWithProgress error: " & Err.Description
    ShowToastWithProgress = False
End Function

'=========================
' UNIFIED MESSAGE BOX
'=========================
Public Function ShowMsgBoxUnified( _
    ByVal Message As String, _
    Optional ByVal Title As String = "Notification", _
    Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
    Optional ByVal Mode As String = "auto", _
    Optional ByVal TimeoutSeconds As Long = 5, _
    Optional ByVal Level As String = "INFO", _
    Optional ByVal LinkUrl As String = "", _
    Optional ByVal CallbackName As String = "", _
    Optional ByVal Icon As String = "", _
    Optional ByVal NoDismiss As Boolean = False, _
    Optional ByVal Sound As String = "", _
    Optional ByVal ImagePath As String = "", _
    Optional ByVal ImageSize As String = "Small", _
    Optional ByVal Position As String = "BR", _
    Optional ByVal Progress As Long = 0) As VbMsgBoxResult
    
    On Error GoTo ErrHandler
    
    If Not MsgBoxMain.IsToastSystemReady Then
        Logs.LogWarn "[MsgBoxUniversal] Toast system not ready; falling back to MsgBox"
        ShowMsgBoxUnified = MsgBox(Message, buttons, Title)
        Exit Function
    End If
    
    Dim t As New clsToastNotification
    With t
        .Title = Title
        .Message = Message
        .Level = Level
        .Duration = TimeoutSeconds
        .Position = Position
        .Icon = Icon
        .SoundName = Sound
        .ImagePath = ImagePath
        .CallbackMacro = CallbackName
        .Progress = Progress
        .YOffset = 0
    End With
    t.Show "auto"
    ShowMsgBoxUnified = vbOK
    Logs.LogInfo "[MsgBoxUniversal] ShowMsgBoxUnified: " & Title & ", Mode: " & Mode
    Exit Function

ErrHandler:
    Logs.LogError "[MsgBoxUniversal] ShowMsgBoxUnified error: " & Err.Description
    ShowMsgBoxUnified = vbCancel
End Function

'=========================
' HELPER FUNCTIONS
'=========================
Public Function EscapeJson(ByVal s As String) As String
    On Error Resume Next
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJson = s
End Function

Public Function EscapeHTML(ByVal Text As String) As String
    On Error Resume Next
    Dim result As String
    result = Text
    result = Replace(result, "&", "&amp;")
    result = Replace(result, "<", "&lt;")
    result = Replace(result, ">", "&gt;")
    result = Replace(result, """", "&quot;")
    result = Replace(result, "'", "&#39;")
    EscapeHTML = result
End Function

Public Function GetTempPath() As String
    GetTempPath = Setup.GetTempFolder()
End Function

'=========================
' DIAGNOSTICS
'=========================
Public Sub DiagnoseToastSystem()
    Dim msg As String
    msg = "Toast System Diagnostics" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    
    ' Check PowerShell listener (KPopListener)
    msg = msg & "PowerShell KPopListener:" & vbCrLf
    If PowershellListenerRunning() Then
        msg = msg & "  ? RUNNING (named pipe available)" & vbCrLf
    Else
        msg = msg & "  ? NOT RUNNING (named pipe unavailable)" & vbCrLf
    End If
    msg = msg & vbCrLf
    
    ' Check Python listener (KPopListener)
    msg = msg & "Python KPopListener:" & vbCrLf
    If PythonListenerRunning() Then
        msg = msg & "  ? RUNNING" & vbCrLf
    Else
        msg = msg & "  ? NOT RUNNING" & vbCrLf
    End If
    msg = msg & vbCrLf
    
    ' Check temp folder
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    msg = msg & "Temp Folder:" & vbCrLf
    msg = msg & "  " & Setup.GetTempFolder() & vbCrLf
    msg = msg & "  Exists: " & fso.FolderExists(Setup.GetTempFolder()) & vbCrLf
    msg = msg & vbCrLf
    
    ' Check named pipe
    msg = msg & "Named Pipe:" & vbCrLf
    msg = msg & "  " & PIPE_PATH & vbCrLf
    msg = msg & vbCrLf
    
    ' Check toast system ready
    msg = msg & "Toast System Ready: " & MsgBoxMain.IsToastSystemReady & vbCrLf
    
    MsgBox msg, vbInformation, "Toast System Diagnostics"
End Sub

'=========================
' TEST UTILITIES
'=========================
Public Sub TestUtilities()
    On Error GoTo ErrHandler
    
    ' Test escaping functions
    Dim testStr As String
    testStr = "Test & <script>alert('Hello')</script> """
    Debug.Print "Original: " & testStr
    Debug.Print "Escaped HTML: " & EscapeHTML(testStr)
    Debug.Print "Escaped JSON: " & EscapeJson(testStr)
    
    ' Diagnose system
    DiagnoseToastSystem
    
    ' Check if listeners are running
    If Not PowershellListenerRunning() And Not PythonListenerRunning() Then
        MsgBox "No KPopListener running!" & vbCrLf & vbCrLf & _
               "Start KPopListener.ps1 first.", vbExclamation, "Test Utilities"
        Exit Sub
    End If
    
    ' Test quick methods
    MsgBox "Testing quick toast methods..." & vbCrLf & vbCrLf & _
           "You should see 4 toasts appear.", vbInformation, "Test Utilities"
    
    MsgInfoEx "Quick info toast test!", "BR"
    Sleep 2000
    
    MsgWarnEx "Quick warning toast test!", "TR"
    Sleep 2000
    
    MsgErrorEx "Quick error toast test!", "TL"
    Sleep 2000
    
    MsgSuccessEx "Quick success toast test!", "BR"
    Sleep 2000
    
    ' Test progress toast
    ShowToastWithProgress "Progress Test", "Processing data", 75, "INFO", 5
    
    MsgBox "Utilities test complete! Check logs for details.", vbInformation, "Test Utilities"
    Logs.LogInfo "[MsgBoxUniversal] TestUtilities completed"
    Exit Sub

ErrHandler:
    Logs.LogError "[MsgBoxUniversal] TestUtilities error: " & Err.Description
    MsgBox "Error in test: " & Err.Description, vbCritical
End Sub
