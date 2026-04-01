Attribute VB_Name = "MsgBoxMain"
'***************************************************************
' Module: MsgBoxMain
' Version: 11.15
' Purpose: Initialize toast notification system and verify listener services
' Author: Keith Swerling + Claude
' Dependencies: Logs.bas, Setup.bas
' Description:
'   This module manages the toast notification system by checking if
'   supporting services (PowerShell/Python listeners) are available.
'   Uses file-based verification instead of process enumeration.
' Changes:
'   - v11.14: Refactored to use file-based service detection (AV-safe)
'   - Removed direct process/path checking that triggers AV
'   - Uses sentinel files and response files for service verification
'   - Added timeout-based health checks
' Updated: 2025-10-27
'***************************************************************
Option Explicit

'=========================
' CONSTANTS
'=========================
Private Const SERVICE_CHECK_TIMEOUT As Long = 3  ' seconds
Private Const SENTINEL_MAX_AGE As Long = 10      ' seconds

'=========================
' SYSTEM INITIALIZATION
'=========================

' Initialize the toast notification system
' This prepares the environment for displaying user notifications
Public Sub SetupToastSystem()
    On Error GoTo ErrorHandler
    
    ' Load configuration from Setup module
    Setup.InitializeSetup
    
    ' Check if automatic service startup is enabled
    If Not Setup.AutoStartToastServers Then
        Logs.DebugLog "[MsgBoxMain] Auto-start disabled. Manual service startup required.", "INFO"
        Exit Sub
    End If
    
    Logs.DebugLog "[MsgBoxMain] Toast system initialized successfully", "INFO"
    Exit Sub

ErrorHandler:
    Logs.DebugLog "[MsgBoxMain] Initialization error: " & Err.Description, "ERROR"
End Sub

'=========================
' SERVICE STATUS VERIFICATION (FILE-BASED)
'=========================

' Check if the toast notification services are available
' Uses file-based detection instead of process enumeration (AV-safe)
' Returns True if at least one listener service is responding
Public Function IsToastSystemReady() As Boolean
    On Error GoTo ErrorHandler
    
    Dim powershellActive As Boolean
    
    ' Only check PowerShell listener (since Python path causes AV issues)
    powershellActive = CheckListenerHealth("PowerShell")
    
    IsToastSystemReady = powershellActive
    
    Logs.DebugLog "[MsgBoxMain] Service status - PowerShell: " & powershellActive, "INFO"
    Exit Function

ErrorHandler:
    Logs.DebugLog "[MsgBoxMain] Service check failed: " & Err.Description, "ERROR"
    IsToastSystemReady = False
End Function

'=========================
' FILE-BASED HEALTH CHECK
'=========================

' Check listener health using sentinel files (no process enumeration)
' This approach is AV-safe as it only checks file timestamps
Private Function CheckListenerHealth(ByVal serviceType As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim sentinelPath As String
    sentinelPath = fso.BuildPath(Setup.GetTempFolder(), "ToastWatcher_Alive.txt")
    
    ' Check if sentinel file exists
    If Not fso.FileExists(sentinelPath) Then
        Logs.DebugLog "[MsgBoxMain] Sentinel file not found: " & sentinelPath, "WARN"
        CheckListenerHealth = False
        Exit Function
    End If
    
    ' Check sentinel file age
    Dim sentinelFile As Object
    Set sentinelFile = fso.GetFile(sentinelPath)
    
    Dim fileAge As Long
    fileAge = DateDiff("s", sentinelFile.DateLastModified, Now)
    
    If fileAge > SENTINEL_MAX_AGE Then
        Logs.DebugLog "[MsgBoxMain] Sentinel file is stale (" & fileAge & "s old)", "WARN"
        CheckListenerHealth = False
    Else
        Logs.DebugLog "[MsgBoxMain] " & serviceType & " listener is healthy (sentinel: " & fileAge & "s old)", "INFO"
        CheckListenerHealth = True
    End If
    
    Set sentinelFile = Nothing
    Set fso = Nothing
    Exit Function

ErrorHandler:
    Logs.DebugLog "[MsgBoxMain] Health check error for " & serviceType & ": " & Err.Description, "ERROR"
    CheckListenerHealth = False
End Function

'=========================
' QUERY-BASED SERVICE CHECK
'=========================

' Send an IsRunning query and wait for response (alternative method)
' Uses file-based communication instead of process enumeration
Public Function QueryListenerStatus() As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim TempFolder As String
    TempFolder = Setup.GetTempFolder()
    
    Dim requestPath As String
    Dim statusPath As String
    requestPath = fso.BuildPath(TempFolder, "ToastRequest.json")
    statusPath = fso.BuildPath(TempFolder, "ToastListenerStatus.json")
    
    ' Clean up old status file
    If fso.FileExists(statusPath) Then
        fso.DeleteFile statusPath, True
    End If
    
    ' Send IsRunning query
    Dim queryJson As String
    queryJson = "{""IsRunningQuery"": true}"
    
    Dim fileStream As Object
    Set fileStream = fso.CreateTextFile(requestPath, True, False)
    fileStream.Write queryJson
    fileStream.Close
    Set fileStream = Nothing
    
    ' Wait for response (with timeout)
    Dim startTime As Double
    startTime = Timer
    
    Do While (Timer - startTime) < SERVICE_CHECK_TIMEOUT
        If fso.FileExists(statusPath) Then
            ' Response received
            Dim responseText As String
            responseText = ReadTextFile(statusPath)
            
            ' Remove UTF-8 BOM if present
            If Len(responseText) > 0 Then
                If AscW(Left(responseText, 1)) = &HFEFF Then
                    responseText = Mid(responseText, 2)
                End If
            End If
            
            ' Check for IsRunning (flexible matching)
            If InStr(1, responseText, """IsRunning""", vbTextCompare) > 0 And _
               InStr(1, responseText, "true", vbTextCompare) > 0 Then
                Logs.DebugLog "[MsgBoxMain] Listener responded to query", "INFO"
                QueryListenerStatus = True
                Exit Function
            End If
        End If
        
        ' Small delay before next check
        Sleep 100
    Loop
    
    ' Timeout reached
    Logs.DebugLog "[MsgBoxMain] Listener query timeout", "WARN"
    QueryListenerStatus = False
    
    Set fso = Nothing
    Exit Function

ErrorHandler:
    Logs.DebugLog "[MsgBoxMain] Query status error: " & Err.Description, "ERROR"
    QueryListenerStatus = False
End Function

'=========================
' HELPER FUNCTIONS
'=========================

' Read text file contents safely
Private Function ReadTextFile(ByVal filePath As String) As String
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim fileStream As Object
    Set fileStream = fso.OpenTextFile(filePath, 1, False)  ' 1 = ForReading
    ReadTextFile = fileStream.ReadAll
    fileStream.Close
    
    Set fileStream = Nothing
    Set fso = Nothing
End Function

' Sleep function (for query timeout)
Private Sub Sleep(ByVal milliseconds As Long)
    On Error Resume Next
    Dim endTime As Double
    endTime = Timer + (milliseconds / 1000)
    Do While Timer < endTime
        DoEvents
    Loop
End Sub

'=========================
' DIAGNOSTIC UTILITIES
'=========================

' Display current system status for troubleshooting
Public Sub ShowSystemStatus()
    On Error Resume Next
    
    Dim statusReport As String
    statusReport = "===== Toast Notification System Status =====" & vbCrLf & vbCrLf
    
    ' Check using sentinel file method
    statusReport = statusReport & "PowerShell Listener (Sentinel): " & _
        IIf(CheckListenerHealth("PowerShell"), "ACTIVE", "INACTIVE") & vbCrLf
    
    ' Check using query method
    statusReport = statusReport & "Listener Query Response: " & _
        IIf(QueryListenerStatus(), "ACTIVE", "INACTIVE") & vbCrLf & vbCrLf
    
    statusReport = statusReport & "System Ready: " & _
        IIf(IsToastSystemReady(), "YES", "NO") & vbCrLf
    
    MsgBox statusReport, vbInformation, "System Status"
End Sub

'=========================
' MAINTENANCE FUNCTIONS
'=========================

' Test connectivity to notification services using file-based methods
Public Function TestNotificationServices() As Boolean
    On Error Resume Next
    
    Dim sentinelOK As Boolean
    Dim queryOK As Boolean
    
    ' Test both methods
    sentinelOK = CheckListenerHealth("PowerShell")
    queryOK = QueryListenerStatus()
    
    TestNotificationServices = (sentinelOK Or queryOK)
    
    Logs.DebugLog "[MsgBoxMain] Service test completed. Sentinel: " & sentinelOK & _
                  ", Query: " & queryOK, "INFO"
End Function

'=========================
' QUICK TOAST TEST
'=========================

' Send a simple test toast to verify the system is working
Public Sub SendTestToast()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim requestPath As String
    requestPath = fso.BuildPath(Setup.GetTempFolder(), "ToastRequest.json")
    
    ' Create simple test toast
    Dim testJson As String
    testJson = "{" & _
               """Title"": ""VBA Test Toast""," & _
               """Message"": ""System is working! Time: " & Format(Now, "hh:mm:ss") & """," & _
               """ToastType"": ""SUCCESS""," & _
               """DurationSec"": 4," & _
               """Position"": ""BR""," & _
               """Icon"": ""?""" & _
               "}"
    
    ' Write to file
    Dim ts As Object
    Set ts = fso.CreateTextFile(requestPath, True, False)
    ts.Write testJson
    ts.Close
    
    Logs.DebugLog "[MsgBoxMain] Test toast sent successfully", "INFO"
    Exit Sub
    
ErrorHandler:
    Logs.DebugLog "[MsgBoxMain] Test toast failed: " & Err.Description, "ERROR"
End Sub

