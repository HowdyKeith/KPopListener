Attribute VB_Name = "MsgBoxUI_Declares"
' =====================================================================
' Module: MsgBoxUI_Declares
' Version: 5.6
' Purpose: API Declarations and Utility Support for MsgBoxUI / ToastWatcher
' Author: Keith Swerling + ChatGPT, Grok, and Claude
' Dependencies: Logs.bas (v1.0.1), Setup.bas (v1.4), MsgBoxUniversal.bas (v3.3), MsgBoxUI.bas (v5.10)
' Changes:
'   - Added declarations for MsgBoxUI_UsePowerShellToasts, MsgBoxUI_UseTempJsonFallback, MsgBoxUI_ToastPipeName
'   - Enhanced logging with Logs.DebugLog
'   - Updated version to 5.6
' Updated: 2025-10-24
' =====================================================================
Option Explicit

' ============================================================
' GLOBALS
' ============================================================
Public MsgBoxUI_UsePowerShellToasts As Boolean ' Initialized in MsgBoxUI_Init
Public MsgBoxUI_UseTempJsonFallback As Boolean ' Initialized in MsgBoxUI_Init
Public MsgBoxUI_ToastPipeName As String        ' Initialized in MsgBoxUI_Init
Public Const MSGBOXUI_VERSION As String = "5.10" ' Matches MsgBoxUI.bas version

' ============================================================
' Win32 API DECLARATIONS
' ============================================================

' Handle utilities
Public Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long

' Pipe + file I/O
Public Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileW" ( _
    ByVal lpFileName As LongPtr, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As LongPtr, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As LongPtr) As LongPtr

Public Declare PtrSafe Function WriteFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As LongPtr) As Long

Public Declare PtrSafe Function ReadFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As LongPtr) As Long

' Process control
Public Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As LongPtr

Public Declare PtrSafe Function TerminateProcess Lib "kernel32" ( _
    ByVal hProcess As LongPtr, _
    ByVal uExitCode As Long) As Long

Public Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

' ============================================================
' CONSTANTS
' ============================================================

Public Const GENERIC_READ As Long = &H80000000
Public Const GENERIC_WRITE As Long = &H40000000
Public Const FILE_SHARE_READ As Long = &H1
Public Const FILE_SHARE_WRITE As Long = &H2
Public Const OPEN_EXISTING As Long = 3
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const INVALID_HANDLE_VALUE As LongPtr = -1

Public Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000
Public Const PROCESS_TERMINATE As Long = &H1

' ============================================================
' UTILITIES
' ============================================================

' === Pipe existence check ===
Public Function PipeExists(ByVal pipeName As String) As Boolean
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    PipeExists = fso.FileExists(pipeName)
    If Err.Number <> 0 Then
        Logs.DebugLog "[MsgBoxUI_Declares] PipeExists error for " & pipeName & ": " & Err.Description, "ERROR"
        PipeExists = False
    End If
End Function

' === JSON fallback existence ===
Public Function ToastJsonPending() As Boolean
    On Error Resume Next
    Dim tmpFile As String
    tmpFile = Setup.TEMP_FOLDER & "\ExcelToasts\ToastRequest.json"
    ToastJsonPending = (Dir(tmpFile) <> "")
    If Err.Number <> 0 Then
        Logs.DebugLog "[MsgBoxUI_Declares] ToastJsonPending error for " & tmpFile & ": " & Err.Description, "ERROR"
        ToastJsonPending = False
    End If
End Function

' === Detect if ToastWatcherRT.ps1 is active ===
Public Function IsToastWatcherRunning() As Boolean
    On Error Resume Next
    Dim wmi As Object, procs As Object, p As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set procs = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='powershell.exe'")
    
    ' Try multiple common locations
    Dim paths As Variant
    paths = Array( _
        Environ$("USERPROFILE") & "\Documents\PowerShell\ToastWatcherRT.ps1", _
        Environ$("USERPROFILE") & "\OneDrive\Documents\PowerShell\ToastWatcherRT.ps1", _
        Environ$("APPDATA") & "\PowerShell\ToastWatcherRT.ps1", _
        Environ$("USERPROFILE") & "\OneDrive\Documents\2025\PowerShell\ToastWatcherRT.ps1" _
    )
    
    Dim searchPath As Variant
    For Each p In procs
        For Each searchPath In paths
            If InStr(LCase$(p.CommandLine), LCase$(CStr(searchPath))) > 0 Then
                IsToastWatcherRunning = True
                Logs.DebugLog "[MsgBoxUI_Declares] ToastWatcherRT found at " & searchPath, "INFO"
                Exit Function
            End If
        Next
    Next
    IsToastWatcherRunning = False
    Logs.DebugLog "[MsgBoxUI_Declares] ToastWatcherRT not running", "INFO"
End Function

' === Attempt graceful termination of ToastWatcherRT ===
Public Function KillToastWatcher() As Boolean
    On Error Resume Next
    Dim wmi As Object, procs As Object, p As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set procs = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='powershell.exe'")
    Dim psPath As String
    psPath = LCase$(Environ$("USERPROFILE") & "\OneDrive\Documents\2025\PowerShell\ToastWatcherRT.ps1")
    Dim found As Boolean
    For Each p In procs
        If InStr(LCase$(p.CommandLine), psPath) > 0 Then
            p.Terminate
            found = True
            Logs.DebugLog "[MsgBoxUI_Declares] Terminated ToastWatcherRT at " & psPath, "INFO"
        End If
    Next
    KillToastWatcher = found
    If Not found Then
        Logs.DebugLog "[MsgBoxUI_Declares] No ToastWatcherRT process found at " & psPath, "INFO"
    End If
End Function

' === Quick diagnostic output ===
Public Sub DebugToastStatus()
    On Error Resume Next
    Dim Status As String
    Status = "=== MsgBoxUI/ToastWatcher Diagnostic ===" & vbCrLf & _
             "MsgBoxUI Version: " & MSGBOXUI_VERSION & vbCrLf & _
             "PowerShell Toasts Enabled: " & MsgBoxUI_UsePowerShellToasts & vbCrLf & _
             "Temp JSON Fallback Enabled: " & MsgBoxUI_UseTempJsonFallback & vbCrLf & _
             "Pipe Name: " & MsgBoxUI_ToastPipeName & vbCrLf & _
             "Listener Running: " & IsToastWatcherRunning & vbCrLf & _
             "Pending Temp JSON: " & ToastJsonPending & vbCrLf & _
             "========================================"
    MsgBox Status, vbInformation, "MsgBoxUI Diagnostic"
    Logs.DebugLog "[MsgBoxUI_Declares] DebugToastStatus displayed: " & Status, "INFO"
End Sub

' === Internal low-level pipe write ===
Public Function WritePipeMessage(ByVal pipeName As String, ByVal Text As String) As Boolean
    On Error Resume Next
    Dim hPipe As LongPtr
    Dim written As Long
    Dim bytes() As Byte
    bytes = StrConv(Text, vbFromUnicode)
    
    hPipe = CreateFile(StrPtr(pipeName), GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hPipe <> INVALID_HANDLE_VALUE Then
        Call WriteFile(hPipe, bytes(0), UBound(bytes) + 1, written, 0)
        CloseHandle hPipe
        WritePipeMessage = (written > 0)
        Logs.DebugLog "[MsgBoxUI_Declares] WritePipeMessage succeeded for " & pipeName & ": " & Text, "INFO"
    Else
        WritePipeMessage = False
        Logs.DebugLog "[MsgBoxUI_Declares] WritePipeMessage failed for " & pipeName & ": " & Err.Description, "ERROR"
    End If
End Function

' ============================================================
' CONSOLE HELPERS (for PowerShell debug)
' ============================================================
Public Sub ConsoleLog(ByVal Text As String)
    On Error Resume Next
    Debug.Print "[MsgBoxUI] " & Text
    Logs.DebugLog "[MsgBoxUI_Declares] ConsoleLog: " & Text, "INFO"
End Sub




