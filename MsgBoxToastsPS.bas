Attribute VB_Name = "MsgBoxToastsPS"
'***************************************************************
' Module: MsgBoxToastsPS.bas
' Version: 6.1
' Purpose: PowerShell Toast Bridge (listener, single-use, progress fallback)
' Features:
'  - ShowToastPowerShell (single-use or listener)
'  - Progress toast via VBS/HTA fallback when listener not present
'  - Embedded helper to build single-use PS1 and VBS for progress
'***************************************************************
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Basic helpers
Private Function tempPath() As String
    tempPath = Environ$("TEMP")
End Function

Private Function EscapeJson(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJson = s
End Function

' Public: ShowToastPowerShell
Public Function ShowToastPowerShell( _
    ByVal Title As String, _
    ByVal Message As String, _
    ByVal durationSec As Long, _
    ByVal ToastType As String, _
    Optional ByVal LinkUrl As String = "", _
    Optional ByVal Icon As String = "", _
    Optional ByVal Sound As String = "", _
    Optional ByVal ImagePath As String = "", _
    Optional ByVal ImageSize As String = "Small", _
    Optional ByVal CallbackMacro As String = "", _
    Optional ByVal NoDismiss As Boolean = False, _
    Optional ByVal Position As String = "BR", _
    Optional ByVal Progress As Long = 0, _
    Optional ByVal SingleUse As Boolean = False, _
    Optional ByVal ProgressFile As String = "") As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim tmp As String: tmp = tempPath()
    Dim requestFile As String: requestFile = tmp & "\ToastRequest.json"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Dim json As String
    
    ' Remove older request files to avoid confusion
    On Error Resume Next
    If fso.FileExists(requestFile) Then fso.DeleteFile requestFile
    On Error GoTo ErrorHandler
    
    ' Build JSON
    json = "{"
    json = json & """Title"":""" & EscapeJson(Title) & ""","
    json = json & """Message"":""" & EscapeJson(Message) & ""","
    json = json & """DurationSec"":" & durationSec & ","
    json = json & """ToastType"":""" & UCase$(ToastType) & ""","
    json = json & """LinkUrl"":""" & EscapeJson(LinkUrl) & ""","
    json = json & """Icon"":""" & EscapeJson(Icon) & ""","
    json = json & """Sound"":""" & Sound & ""","
    json = json & """ImagePath"":""" & EscapeJson(ImagePath) & ""","
    json = json & """ImageSize"":""" & ImageSize & ""","
    json = json & """CallbackMacro"":""" & CallbackMacro & ""","
    json = json & """NoDismiss"":" & IIf(NoDismiss, "true", "false") & ","
    json = json & """Position"":""" & Position & ""","
    json = json & """Progress"":" & Progress
    If Len(ProgressFile) > 0 Then
        json = json & ",""ProgressFile"":""" & EscapeJson(ProgressFile) & """"
    End If
    json = json & "}"
    
    ' Play beep if requested
    If UCase$(Sound) = "BEEP" Then
        On Error Resume Next
        Beep 800, 200
        On Error GoTo ErrorHandler
    End If
    
    ' If single-use or listener not running, fallback
    If SingleUse Or Not PowershellListenerRunning() Then
        If Progress > 0 Then
            ' Use VBS+HTA fallback for progress UI (persistent)
            Dim vbsFile As String: vbsFile = tmp & "\ProgressToast.vbs"
            Dim htaFile As String: htaFile = tmp & "\ProgressToast.hta"
            Set ts = fso.CreateTextFile(vbsFile, True, True)
            ts.Write GetVBSCode(IIf(Len(ProgressFile) > 0, ProgressFile, tmp & "\ProgressRequest.json"), htaFile)
            ts.Close
            
            ' initialize progress file
            Dim pJson As String
            pJson = "{""Progress"":" & Progress & ",""Message"":""" & EscapeJson(Message) & """,""Running"":true}"
            Set ts = fso.CreateTextFile(IIf(Len(ProgressFile) > 0, ProgressFile, tmp & "\ProgressRequest.json"), True, True)
            ts.Write pJson
            ts.Close
            
            On Error Resume Next
            CreateObject("WScript.Shell").Run "wscript """ & vbsFile & """", 0, False
            If Err.Number <> 0 Then
                Debug.Print "[ShowToastPowerShell] Failed to launch VBS: " & Err.Description
                ShowToastPowerShell = False
                Exit Function
            End If
            ShowToastPowerShell = True
            Exit Function
        Else
            ' Single-use PowerShell toast
            Dim singleUsePs1 As String: singleUsePs1 = tmp & "\SingleUseToast.ps1"
            Dim psCode As String
            psCode = BuildSingleUsePs1Code(Title, Message, durationSec, ToastType, LinkUrl, Icon, Sound, ImagePath, ImageSize, CallbackMacro, NoDismiss, Position, Progress)
            Set ts = fso.CreateTextFile(singleUsePs1, True, True)
            ts.Write psCode
            ts.Close
            
            Dim cmd As String
            cmd = "powershell -NoProfile -ExecutionPolicy Bypass -File """ & singleUsePs1 & """"
            On Error Resume Next
            CreateObject("WScript.Shell").Run cmd, 0, False
            If Err.Number <> 0 Then
                Debug.Print "[ShowToastPowerShell] Failed to start PowerShell single-use: " & Err.Description
                ShowToastPowerShell = False
                Exit Function
            End If
            Sleep 1000
            If fso.FileExists(singleUsePs1) Then fso.DeleteFile singleUsePs1
            ShowToastPowerShell = True
            Exit Function
        End If
    End If
    
    ' Listener mode: write request file and wait for PS to process
    Set ts = fso.CreateTextFile(requestFile, True, True)
    ts.Write json
    ts.Close
    Debug.Print "[ShowToastPowerShell] Wrote request: " & requestFile
    
    Dim maxWait As Long: maxWait = 30
    Dim i As Long
    For i = 1 To maxWait
        If Not fso.FileExists(requestFile) Then
            Debug.Print "[ShowToastPowerShell] Listener processed request."
            ShowToastPowerShell = True
            Exit Function
        End If
        Sleep 100
        DoEvents
    Next i
    
    Debug.Print "[ShowToastPowerShell] Timeout waiting for listener."
    ShowToastPowerShell = False
    Exit Function

ErrorHandler:
    Debug.Print "[ShowToastPowerShell] Error: " & Err.Description
    ShowToastPowerShell = False
End Function

' Build single-use PowerShell script code for toast (BurntToast if available; otherwise HTA)
Private Function BuildSingleUsePs1Code( _
    Title As String, Message As String, durationSec As Long, ToastType As String, _
    LinkUrl As String, Icon As String, Sound As String, ImagePath As String, _
    ImageSize As String, CallbackMacro As String, NoDismiss As Boolean, _
    Position As String, Progress As Long) As String
    
    Dim psCode As String
    psCode = "$ErrorActionPreference = 'Stop'" & vbCrLf
    psCode = psCode & "try {" & vbCrLf
    psCode = psCode & "  if (Get-Module -ListAvailable -Name BurntToast) {" & vbCrLf
    psCode = psCode & "    Import-Module BurntToast" & vbCrLf
    psCode = psCode & "    $params = @{" & vbCrLf
    psCode = psCode & "      Text = @('" & Replace(Title, "'", "''") & "', '" & Replace(Message, "'", "''") & "')" & vbCrLf
    psCode = psCode & "      AppId = 'Microsoft.Office.Excel'" & vbCrLf
    If Len(ImagePath) > 0 Then
        psCode = psCode & "      if (Test-Path '" & Replace(ImagePath, "'", "''") & "') { $params['HeroImage'] = '" & Replace(ImagePath, "'", "''") & "' }" & vbCrLf
    End If
    If Len(LinkUrl) > 0 Then
        psCode = psCode & "      $actions = New-BTAction -Content 'Open' -Arguments '" & Replace(LinkUrl, "'", "''") & "'" & vbCrLf
    End If
    If durationSec > 0 Then
        psCode = psCode & "      $params['ExpirationTime'] = (Get-Date).AddSeconds(" & durationSec & ")" & vbCrLf
    End If
    psCode = psCode & "    }" & vbCrLf
    psCode = psCode & "    New-BurntToastNotification @params" & vbCrLf
    psCode = psCode & "  } else {" & vbCrLf
    psCode = psCode & "    # Fallback: generate HTA and open via mshta (simple presentation)" & vbCrLf
    psCode = psCode & "    $hta = '<html><body><h3>" & Replace(Title, "'", "''") & "</h3><p>" & Replace(Message, "'", "''") & "</p></body></html>'" & vbCrLf
    psCode = psCode & "    $htaPath = [System.IO.Path]::Combine($env:TEMP,'Toast_'+[System.Guid]::NewGuid().ToString()+'.hta')" & vbCrLf
    psCode = psCode & "    Set-Content -Path $htaPath -Value $hta -Encoding UTF8" & vbCrLf
    psCode = psCode & "    Start-Process 'mshta.exe' -ArgumentList ('""" & "$htaPath" & """') -WindowStyle Hidden" & vbCrLf
    psCode = psCode & "  }" & vbCrLf
    psCode = psCode & "} catch { $_.Exception.Message | Out-File -FilePath ([System.IO.Path]::Combine($env:TEMP,'ToastResponse.txt')) -Encoding utf8 }" & vbCrLf
    
    BuildSingleUsePs1Code = psCode
End Function

' VBS code generator for progress HTA (used when listener not present)
Private Function GetVBSCode(ProgressFile As String, htaFile As String) As String
    Dim vbsCode As String
    vbsCode = "Option Explicit" & vbCrLf
    vbsCode = vbsCode & "Dim fso, shell" & vbCrLf
    vbsCode = vbsCode & "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    vbsCode = vbsCode & "Set shell = CreateObject(""WScript.Shell"")" & vbCrLf
    vbsCode = vbsCode & "RunDynamicProgressToast ""Progress"", ""Processing..."", ""INFO"", """", """", """", 0, ""C"", 0, """ & Replace(ProgressFile, """", """""") & """, """ & Replace(htaFile, """", """""") & """" & vbCrLf
    vbsCode = vbsCode & GetRunDynamicProgressToastCode()
    GetVBSCode = vbsCode
End Function

' This returns VBScript that will create and run an HTA reading the given ProgressFile
Private Function GetRunDynamicProgressToastCode() As String
    Dim code As String
    code = "Sub RunDynamicProgressToast(Title, Message, ToastType, LinkUrl, ImagePath, CallbackMacro, DurationSec, Position, Progress, ProgressFile, HTAFile)" & vbCrLf
    code = code & "  Dim fso" & vbCrLf
    code = code & "  Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    code = code & "  Dim html" & vbCrLf
    code = code & "  html = ""<html><head><meta charset='utf-8'><title>"" & Title & ""</title></head><body>"" & Message & ""</body></html>""" & vbCrLf
    code = code & "  Dim ts: Set ts = fso.CreateTextFile(HTAFile, True): ts.Write html: ts.Close" & vbCrLf
    code = code & "  CreateObject(""WScript.Shell"").Run ""mshta.exe """" & HTAFile & """""", 0, False" & vbCrLf
    code = code & "End Sub" & vbCrLf
    GetRunDynamicProgressToastCode = code
End Function

' Check if PowerShell listener sentinel file is present and recent
Public Function PowershellListenerRunning() As Boolean
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tmp As String: tmp = tempPath()
    Dim sentinel As String: sentinel = tmp & "\ToastLastProcessed.txt"
    If fso.FileExists(sentinel) Then
        Dim ts As Object: Set ts = fso.OpenTextFile(sentinel, 1)
        Dim lastTimeStr As String: lastTimeStr = ts.ReadLine
        ts.Close
        Dim diff As Double
        diff = DateDiff("s", CDate(lastTimeStr), Now)
        PowershellListenerRunning = (diff < 10)
    Else
        PowershellListenerRunning = False
    End If
    On Error GoTo 0
End Function

' Utility: small cross-platform beep (uses PowerShell)
Public Sub Beep(Frequency As Long, Duration As Long)
    On Error Resume Next
    CreateObject("WScript.Shell").Run "powershell -Command [Console]::Beep(" & Frequency & "," & Duration & ")", 0, True
End Sub


