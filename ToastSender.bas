Attribute VB_Name = "ToastSender"
Option Explicit
' ============================================================
' Module: ToastSender
' Version: 5.5
' Purpose: Unified VBA ? KPopListener Toast Request Sender
' Works with: KPopListener.ps1 v6.8b+ and KPopListener.py v7.2+
' Author: Keith Swerling + Claude
' Dependencies: Setup.bas (v1.8), Logs.bas (v1.0.3)
' Changes:
'   - v5.5: Updated for KPopListener naming (pipe, MMF, temp folder)
'   - v5.4: Fixed named pipe, MMF, and temp file communication
'   - v5.4: Added proper error handling and logging
'   - v5.4: Fixed JSON escaping and array handling
' Features:
'   - Sends JSON toast payloads to listener via:
'       • Named Pipe ("Pipe") - Direct communication
'       • Memory-Mapped File ("MMF") - Shared memory
'       • Temp file fallback ("Temp") - File-based
'       • Direct PS1 execution ("SendMode") - Bypass listener
'   - Supports Unicode titles/messages, buttons, menus
' Updated: 2025-10-29
' ============================================================

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
    ' Named Pipe API declarations
    Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
        ByVal lpFileName As String, _
        ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, _
        ByVal lpSecurityAttributes As Long, _
        ByVal dwCreationDisposition As Long, _
        ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) As Long
    
    Private Declare PtrSafe Function WriteFile Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByVal lpBuffer As String, _
        ByVal nNumberOfBytesToWrite As Long, _
        ByRef lpNumberOfBytesWritten As Long, _
        ByVal lpOverlapped As Long) As Long
    
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
    Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
        ByVal lpFileName As String, _
        ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, _
        ByVal lpSecurityAttributes As Long, _
        ByVal dwCreationDisposition As Long, _
        ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) As Long
    
    Private Declare Function WriteFile Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByVal lpBuffer As String, _
        ByVal nNumberOfBytesToWrite As Long, _
        ByRef lpNumberOfBytesWritten As Long, _
        ByVal lpOverlapped As Long) As Long
    
    Private Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
#End If

' Constants for CreateFile
Private Const GENERIC_WRITE As Long = &H40000000
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const INVALID_HANDLE_VALUE As Long = -1

' KPopListener Configuration
Private Const PIPE_PATH As String = "\\.\pipe\ExcelToastPipe"
Private Const MMF_NAME As String = "KPopListenerMMF"
Private Const TEMP_JSON As String = "ToastRequest.json"

' ------------------ Public Interface ------------------
Public Sub SendToast( _
    ByVal Title As String, _
    ByVal Message As String, _
    Optional ByVal ToastType As String = "INFO", _
    Optional ByVal durationSec As Long = 5, _
    Optional ByVal buttons As Variant, _
    Optional ByVal menu As Variant, _
    Optional ByVal Mode As String = "Temp")

    On Error GoTo ErrorHandler
    
    Logs.LogDebug "[ToastSender] Sending toast: " & Title
    
    Dim toastJson As String, psCmd As String
    Dim btnJson As String, menuJson As String
    
    ' Validate Mode
    If Len(Mode) = 0 Then Mode = "Temp"

    ' --- Build Button JSON ---
    btnJson = "[]"
    If Not IsMissing(buttons) Then
        If IsArray(buttons) Then
            btnJson = "["
            Dim b As Variant
            For Each b In buttons
                If Len(btnJson) > 1 Then btnJson = btnJson & ","
                btnJson = btnJson & """" & EscapeJson(CStr(b)) & """"
            Next b
            btnJson = btnJson & "]"
        End If
    End If

    ' --- Build Menu JSON ---
    menuJson = "[]"
    If Not IsMissing(menu) Then
        If IsArray(menu) Then
            menuJson = "["
            Dim m As Variant
            For Each m In menu
                If Len(menuJson) > 1 Then menuJson = menuJson & ","
                menuJson = menuJson & """" & EscapeJson(CStr(m)) & """"
            Next m
            menuJson = menuJson & "]"
        End If
    End If

    ' --- Build Payload ---
    toastJson = "{""Title"":""" & EscapeJson(Title) & """," & _
                 """Message"":""" & EscapeJson(Message) & """," & _
                 """ToastType"":""" & ToastType & """," & _
                 """DurationSec"":" & durationSec & "," & _
                 """Buttons"":" & btnJson & "," & _
                 """Menu"":" & menuJson & "}"
    
    Logs.LogDebug "[ToastSender] JSON: " & Left(toastJson, 100) & "..."
    Logs.LogDebug "[ToastSender] Mode: " & Mode

    ' --- Send via Mode ---
    Dim success As Boolean
    success = False
    
    Select Case LCase(Mode)
        Case "sendmode"
            ' Direct PowerShell execution
            Logs.LogInfo "[ToastSender] Using SendMode (direct PS1 execution)"
            psCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & _
                    GetListenerPath() & """ -SendToastJson '" & Replace(toastJson, "'", "''") & "'"
            shell psCmd, vbHide
            success = True
            Logs.LogInfo "[ToastSender] SendMode command executed"

        Case "pipe"
            Logs.LogInfo "[ToastSender] Attempting Named Pipe..."
            success = WriteToNamedPipe(toastJson)
            If Not success Then
                Logs.LogWarn "[ToastSender] Named Pipe failed, falling back to Temp"
                success = WriteJsonFallback(TEMP_JSON, toastJson)
            End If

        Case "mmf"
            Logs.LogInfo "[ToastSender] Attempting Memory-Mapped File..."
            success = WriteToMMF(toastJson)
            If Not success Then
                Logs.LogWarn "[ToastSender] MMF failed, falling back to Temp"
                success = WriteJsonFallback(TEMP_JSON, toastJson)
            End If

        Case "temp"
            Logs.LogInfo "[ToastSender] Using Temp file mode"
            success = WriteJsonFallback(TEMP_JSON, toastJson)

        Case Else
            Logs.LogError "[ToastSender] Invalid Toast Mode: " & Mode
            MsgBox "Invalid Toast Mode: " & Mode, vbCritical, "ToastSender"
            Exit Sub
    End Select
    
    If success Then
        Logs.LogInfo "[ToastSender] Toast sent successfully via " & Mode
    Else
        Logs.LogError "[ToastSender] Failed to send toast via " & Mode
    End If
    
    Exit Sub

ErrorHandler:
    Logs.LogError "[ToastSender] Error in SendToast: " & Err.Description
    MsgBox "Error sending toast: " & Err.Description, vbCritical, "ToastSender"
End Sub

' ============================================================
' =============== COMMUNICATION ROUTINES =====================
' ============================================================

' --- Write JSON to Named Pipe (using Windows API) ---
Private Function WriteToNamedPipe(ByVal json As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim hPipe As Long
    Dim bytesWritten As Long
    Dim jsonBytes As String
    
    ' Open the named pipe
    hPipe = CreateFile(PIPE_PATH, GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    
    If hPipe = INVALID_HANDLE_VALUE Then
        Logs.LogWarn "[ToastSender] Named pipe not available (KPopListener not running?)"
        WriteToNamedPipe = False
        Exit Function
    End If
    
    ' Write JSON to pipe
    jsonBytes = json & vbNullChar
    If WriteFile(hPipe, jsonBytes, Len(jsonBytes), bytesWritten, 0) = 0 Then
        Logs.LogError "[ToastSender] Failed to write to named pipe"
        CloseHandle hPipe
        WriteToNamedPipe = False
        Exit Function
    End If
    
    ' Close pipe handle
    CloseHandle hPipe
    
    Logs.LogInfo "[ToastSender] Written to named pipe: " & bytesWritten & " bytes"
    WriteToNamedPipe = True
    Exit Function
    
ErrorHandler:
    If hPipe <> 0 And hPipe <> INVALID_HANDLE_VALUE Then CloseHandle hPipe
    Logs.LogError "[ToastSender] Named pipe error: " & Err.Description
    WriteToNamedPipe = False
End Function

' --- Write JSON to Memory-Mapped File (via PowerShell helper) ---
Private Function WriteToMMF(ByVal json As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim psCmd As String
    Dim escapedJson As String
    
    ' Escape JSON for PowerShell (replace single quotes)
    escapedJson = Replace(json, "'", "''")
    escapedJson = Replace(escapedJson, """", "'")
    
    ' Build PowerShell command to write to MMF
    psCmd = "powershell.exe -NoProfile -WindowStyle Hidden -Command """ & _
            "$mmf=[System.IO.MemoryMappedFiles.MemoryMappedFile]::CreateOrOpen('" & MMF_NAME & "',65536);" & _
            "$s=$mmf.CreateViewStream();" & _
            "$w=New-Object IO.StreamWriter($s);" & _
            "$w.Write('" & escapedJson & "');" & _
            "$w.Flush();" & _
            "$w.Close();" & _
            "$s.Close();" & _
            "$mmf.Dispose();"""
    
    Logs.LogDebug "[ToastSender] MMF command: " & Left(psCmd, 200) & "..."
    
    shell psCmd, vbHide
    Sleep 500 ' Give PowerShell time to write
    
    Logs.LogInfo "[ToastSender] Written to MMF"
    WriteToMMF = True
    Exit Function
    
ErrorHandler:
    Logs.LogError "[ToastSender] MMF error: " & Err.Description
    WriteToMMF = False
End Function

' --- JSON fallback file (TEMP) ---
Private Function WriteJsonFallback(ByVal fileName As String, ByVal json As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim tmpPath As String
    tmpPath = Setup.GetTempFolder() & "\" & fileName
    
    ' Write JSON to temp file
    Dim ts As Object
    Set ts = fso.CreateTextFile(tmpPath, True, True) ' Unicode = True
    ts.Write json
    ts.Close
    
    Logs.LogInfo "[ToastSender] Written to temp file: " & tmpPath
    WriteJsonFallback = True
    Exit Function
    
ErrorHandler:
    Logs.LogError "[ToastSender] Temp file error: " & Err.Description
    WriteJsonFallback = False
End Function

' ============================================================
' ===================== UTILITIES ============================
' ============================================================

Private Function EscapeJson(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJson = s
End Function

Private Function GetListenerPath() As String
    ' Try to find KPopListener in common locations
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check OneDrive location first
    Dim path As String
    path = "C:\Users\howdy\OneDrive\MsgBox\KPopListener.ps1"
    If fso.FileExists(path) Then
        GetListenerPath = path
        Exit Function
    End If
    
    ' Check temp folder
    path = Setup.GetTempFolder() & "\KPopListener.ps1"
    If fso.FileExists(path) Then
        GetListenerPath = path
        Exit Function
    End If
    
    ' Check same folder as workbook
    path = ThisWorkbook.path & "\KPopListener.ps1"
    If fso.FileExists(path) Then
        GetListenerPath = path
        Exit Function
    End If
    
    ' Default to OneDrive location
    GetListenerPath = "C:\Users\howdy\OneDrive\MsgBox\KPopListener.ps1"
    Logs.LogWarn "[ToastSender] KPopListener not found, using default: " & GetListenerPath
End Function

' ============================================================
' ===================== TEST ROUTINES ========================
' ============================================================

Public Sub TestToastPipe()
    Logs.LogInfo "[ToastSender] ===== Testing Named Pipe ====="
    SendToast "Pipe Test", "Sent through Named Pipe at " & Now, "INFO", 5, _
              Array("OK", "Cancel"), Array("Option1", "Option2"), "Pipe"
End Sub

Public Sub TestToastMMF()
    Logs.LogInfo "[ToastSender] ===== Testing Memory-Mapped File ====="
    SendToast "MMF Test", "Sent through Memory-Mapped File at " & Now, "SUCCESS", 5, _
              Array("Close"), Array("A", "B"), "MMF"
End Sub

Public Sub TestToastFallback()
    Logs.LogInfo "[ToastSender] ===== Testing Temp Fallback ====="
    SendToast "Temp Fallback", "Sent through Temp JSON file at " & Now, "WARN", 5, , , "Temp"
End Sub

Public Sub TestToastSendMode()
    Logs.LogInfo "[ToastSender] ===== Testing SendMode (Direct PS1) ====="
    SendToast "SendMode Test", "Sent by direct PS1 execution at " & Now, "INFO", 5, , , "SendMode"
End Sub

Public Sub TestAllModes()
    Logs.LogInfo "[ToastSender] ===== Testing ALL Toast Modes ====="
    
    MsgBox "Testing all toast modes. Check your logs!" & vbCrLf & vbCrLf & _
           "Modes: Temp ? Pipe ? MMF ? SendMode", vbInformation, "Toast Test"
    
    ' Test 1: Temp (most reliable)
    SendToast "Test 1/4: Temp File", "Using temp file method", "INFO", 3, , , "Temp"
    Sleep 3500
    
    ' Test 2: Named Pipe
    SendToast "Test 2/4: Named Pipe", "Using named pipe method", "SUCCESS", 3, , , "Pipe"
    Sleep 3500
    
    ' Test 3: Memory-Mapped File
    SendToast "Test 3/4: MMF", "Using memory-mapped file method", "WARN", 3, , , "MMF"
    Sleep 3500
    
    ' Test 4: Direct PowerShell
    SendToast "Test 4/4: SendMode", "Using direct PS1 execution", "INFO", 3, , , "SendMode"
    
    MsgBox "Test complete! Check logs for results.", vbInformation, "Toast Test"
End Sub

' Start the KPopListener
Public Sub StartToastListener()
    On Error GoTo ErrorHandler
    
    Dim listenerPath As String
    listenerPath = GetListenerPath()
    
    ' Check if file exists
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(listenerPath) Then
        MsgBox "KPopListener.ps1 not found at:" & vbCrLf & listenerPath, vbCritical
        Exit Sub
    End If
    
    ' Check if already running
    Dim hPipe As Long
    hPipe = CreateFile(PIPE_PATH, GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hPipe <> INVALID_HANDLE_VALUE Then
        CloseHandle hPipe
        MsgBox "KPopListener is already running!", vbInformation, "KPopListener"
        Exit Sub
    End If
    
    ' Start PowerShell listener in background
    Dim psCmd As String
    psCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & listenerPath & """"
    
    shell psCmd, vbHide
    
    Logs.LogInfo "[ToastSender] Started KPopListener: " & listenerPath
    
    ' Wait for listener to initialize (max 10 seconds)
    Dim attempts As Integer
    Dim maxAttempts As Integer
    maxAttempts = 20 ' 20 attempts x 500ms = 10 seconds
    
    For attempts = 1 To maxAttempts
        Sleep 500
        
        ' Try to connect to pipe
        hPipe = CreateFile(PIPE_PATH, GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
        
        If hPipe <> INVALID_HANDLE_VALUE Then
            CloseHandle hPipe
            MsgBox "KPopListener started successfully!" & vbCrLf & vbCrLf & _
                   "Took " & (attempts * 0.5) & " seconds to initialize." & vbCrLf & _
                   "Named pipe is now available.", vbInformation, "KPopListener"
            Logs.LogInfo "[ToastSender] KPopListener ready after " & attempts & " attempts"
            Exit Sub
        End If
        
        ' Show progress every 2 seconds
        If attempts Mod 4 = 0 Then
            Logs.LogDebug "[ToastSender] Waiting for KPopListener... (" & attempts & "/" & maxAttempts & ")"
        End If
    Next attempts
    
    ' Timeout
    MsgBox "Listener process started but named pipe did not become available." & vbCrLf & vbCrLf & _
           "Possible issues:" & vbCrLf & _
           "• PowerShell execution policy blocking script" & vbCrLf & _
           "• KPopListener.ps1 has errors" & vbCrLf & _
           "• Named pipe creation failed" & vbCrLf & vbCrLf & _
           "You can still use 'Temp' mode for toasts!", _
           vbExclamation, "KPopListener Timeout"
    
    Logs.LogWarn "[ToastSender] KPopListener timeout after " & maxAttempts & " attempts"
    
    Exit Sub
    
ErrorHandler:
    Logs.LogError "[ToastSender] Error starting KPopListener: " & Err.Description
    MsgBox "Error starting listener: " & Err.Description, vbCritical
End Sub

' Stop the KPopListener
Public Sub StopToastListener()
    On Error Resume Next
    
    ' Send stop command via temp file
    Dim stopJson As String
    stopJson = "{""Command"":""Stop""}"
    
    WriteJsonFallback "ToastStopCommand.json", stopJson
    
    Logs.LogInfo "[ToastSender] Stop command sent to KPopListener"
    MsgBox "Stop command sent to KPopListener.", vbInformation, "KPopListener"
End Sub

Public Sub CheckToastListener()
    Dim msg As String
    msg = "KPopListener Diagnostics" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    
    ' Check listener file
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim listenerPath As String
    listenerPath = GetListenerPath()
    
    msg = msg & "Listener Path:" & vbCrLf
    msg = msg & "  " & listenerPath & vbCrLf
    msg = msg & "  Exists: " & fso.FileExists(listenerPath) & vbCrLf & vbCrLf
    
    ' Check temp folder
    msg = msg & "Temp Folder:" & vbCrLf
    msg = msg & "  " & Setup.GetTempFolder() & vbCrLf
    msg = msg & "  Exists: " & fso.FolderExists(Setup.GetTempFolder()) & vbCrLf & vbCrLf
    
    ' Check for temp JSON file
    Dim tempJsonPath As String
    tempJsonPath = Setup.GetTempFolder() & "\" & TEMP_JSON
    msg = msg & "Temp JSON File:" & vbCrLf
    msg = msg & "  " & tempJsonPath & vbCrLf
    msg = msg & "  Exists: " & fso.FileExists(tempJsonPath) & vbCrLf & vbCrLf
    
    ' Check named pipe (try to connect)
    msg = msg & "Named Pipe Status:" & vbCrLf
    msg = msg & "  Path: " & PIPE_PATH & vbCrLf
    Dim hPipe As Long
    hPipe = CreateFile(PIPE_PATH, GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hPipe = INVALID_HANDLE_VALUE Then
        msg = msg & "  ? NOT AVAILABLE (KPopListener not running?)" & vbCrLf
    Else
        msg = msg & "  ? AVAILABLE (KPopListener is running!)" & vbCrLf
        CloseHandle hPipe
    End If
    
    Logs.LogInfo "[ToastSender] Diagnostics check completed"
    MsgBox msg, vbInformation, "KPopListener Diagnostics"
End Sub

