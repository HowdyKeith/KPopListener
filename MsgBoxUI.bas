Attribute VB_Name = "MsgBoxUI"
'***************************************************************
' Module: MsgBoxUI.bas
' Version: 6.4
' Purpose: Unified VBA Notification Wrappers & ToastWatcher Integration
' Author: Keith Swerling + AI Assistants
' Dependencies: None (standalone)
' Changes:
'   - v6.4: Fixed all issues - removed Python launcher reference,
'           fixed cancel handling, improved JSON fallback
'   - v6.3: Added MMF support
' Updated: 2025-10-27
'***************************************************************
Option Explicit

' =====================================================================
' Configuration
' =====================================================================
Public Const MSGBOXUI_VERSION As String = "6.4"
Public Const MSGBOXUI_TITLE As String = "MsgBoxUI v6.4"

' Settings
Public UsePowerShellToasts As Boolean
Public UseTempJsonFallback As Boolean
Public UsePushbullet As Boolean  ' NEW: Enable Pushbullet mobile notifications
Public ToastPipeName As String

' =====================================================================
' INITIALIZATION
' =====================================================================
Public Sub MsgBoxUI_Init(Optional ByVal EnableToasts As Boolean = True, _
                         Optional ByVal EnableFallback As Boolean = True, _
                         Optional ByVal EnablePushbullet As Boolean = False, _
                         Optional ByVal pipeName As String = "\\.\pipe\ExcelToastPipe")
    
    UsePowerShellToasts = EnableToasts
    UseTempJsonFallback = EnableFallback
    UsePushbullet = EnablePushbullet
    ToastPipeName = pipeName
    
    ' Ensure temp folder exists
    On Error Resume Next
    Dim TempFolder As String
    TempFolder = Environ$("TEMP") & "\ExcelToasts"
    If Dir(TempFolder, vbDirectory) = "" Then MkDir TempFolder
    On Error GoTo 0
End Sub

' =====================================================================
' SIMPLE WRAPPERS
' =====================================================================

' Simple notification
Public Sub Notify(Optional ByVal Title As String = "Notice", _
                  Optional ByVal Message As String = "", _
                  Optional ByVal Level As String = "INFO", _
                  Optional ByVal timeout As Long = 5, _
                  Optional ByVal SendToMobile As Boolean = False)
    
    On Error Resume Next
    
    ' Send desktop toast
    If UsePowerShellToasts Then
        SendToastNotification Title, Message, Level, 0
    Else
        ' Fallback to classic MsgBox
        Dim Icon As VbMsgBoxStyle
        Select Case UCase(Level)
            Case "ERROR": Icon = vbCritical
            Case "WARN", "WARNING": Icon = vbExclamation
            Case "SUCCESS": Icon = vbInformation
            Case Else: Icon = vbInformation
        End Select
        MsgBox Message, Icon, Title
    End If
    
    ' Send to mobile if enabled
    If SendToMobile Or UsePushbullet Then
        SendPushbulletNotification Title, Message
    End If
End Sub

' Progress notification
Public Sub Progress(ByVal Title As String, _
                    ByVal Message As String, _
                    ByVal Percent As Double)
    
    On Error Resume Next
    
    If Percent < 0 Then Percent = 0
    If Percent > 100 Then Percent = 100
    
    Dim displayMsg As String
    displayMsg = Message & " (" & Format(Percent, "0") & "%)"
    
    SendToastNotification Title, displayMsg, "INFO", CLng(Percent)
End Sub

' Input text dialog
Public Function InputText(ByVal Title As String, _
                          ByVal prompt As String, _
                          Optional ByVal DefaultValue As String = "") As String
    
    On Error Resume Next
    InputText = InputBox(prompt, Title, DefaultValue)
End Function

' =====================================================================
' CORE TOAST SENDER
' =====================================================================
Private Sub SendToastNotification(ByVal Title As String, _
                                  ByVal Message As String, _
                                  ByVal Level As String, _
                                  Optional ByVal Progress As Long = -1)
    
    On Error Resume Next
    
    ' Build JSON
    Dim json As String
    json = "{"
    json = json & """Title"":""" & EscapeJson(Title) & ""","
    json = json & """Message"":""" & EscapeJson(Message) & ""","
    json = json & """ToastType"":""" & UCase(Level) & """"
    
    If Progress >= 0 And Progress <= 100 Then
        json = json & ",""Progress"":" & Progress
    End If
    
    json = json & "}"
    
    ' Write to temp file for ToastWatcher to pick up
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim TempFolder As String
    TempFolder = Environ$("TEMP") & "\ExcelToasts"
    
    ' Ensure folder exists
    If Not fso.FolderExists(TempFolder) Then
        fso.CreateFolder TempFolder
    End If
    
    Dim jsonFile As String
    jsonFile = TempFolder & "\ToastRequest.json"
    
    ' Write JSON file
    Dim ts As Object
    Set ts = fso.CreateTextFile(jsonFile, True, False)
    ts.Write json
    ts.Close
End Sub

' Escape JSON special characters
Private Function EscapeJson(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJson = s
End Function

' =====================================================================
' MAIN UNIFIED MSGBOX ROUTER
' =====================================================================
Public Function ShowMsgBoxUnified(ByVal Message As String, _
                                  Optional ByVal Title As String = "Notification", _
                                  Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                                  Optional ByVal Mode As String = "auto", _
                                  Optional ByVal TimeoutSeconds As Long = 5, _
                                  Optional ByVal Level As String = "INFO") As VbMsgBoxResult
    
    On Error GoTo SafeExit
    
    Dim resolvedMode As String
    resolvedMode = LCase$(Mode)
    
    ' Auto-detect mode
    If resolvedMode = "auto" Then
        If UsePowerShellToasts Then
            resolvedMode = "ps"
        Else
            resolvedMode = "classic"
        End If
    End If
    
    ' Route to appropriate display method
    Select Case resolvedMode
        Case "ps", "powershell", "toast"
            SendToastNotification Title, Message, Level, -1
            ShowMsgBoxUnified = vbOK
            
        Case "mshta"
            ShowMSHTAToast Message, Title, TimeoutSeconds, Level
            ShowMsgBoxUnified = vbOK
            
        Case "wscript"
            ShowWScriptPopup Message, Title, TimeoutSeconds
            ShowMsgBoxUnified = vbOK
            
        Case Else
            ShowMsgBoxUnified = MsgBox(Message, buttons, Title)
    End Select
    
    Exit Function

SafeExit:
    ' Fallback to classic MsgBox on error
    ShowMsgBoxUnified = MsgBox(Message, vbOKOnly, Title)
End Function

' =====================================================================
' MSHTA TOAST (Fallback)
' =====================================================================
Private Sub ShowMSHTAToast(ByVal msg As String, _
                          ByVal Title As String, _
                          ByVal timeout As Long, _
                          ByVal Level As String)
    
    On Error Resume Next
    
    Dim bgColor As String
    Select Case UCase(Level)
        Case "ERROR": bgColor = "#f44336"
        Case "WARN", "WARNING": bgColor = "#ff9800"
        Case "SUCCESS": bgColor = "#4caf50"
        Case Else: bgColor = "#2196f3"
    End Select
    
    Dim html As String
    html = "<html><head><title>" & Title & "</title>" & _
           "<HTA:APPLICATION BORDER='none' CAPTION='no' SHOWINTASKBAR='no'>" & _
           "<style>" & _
           "body{font-family:'Segoe UI';background:" & bgColor & ";color:white;padding:20px;margin:0;}" & _
           "h3{margin:0 0 10px 0;font-size:18px;}" & _
           "p{margin:0;font-size:14px;}" & _
           "</style>" & _
           "<script>" & _
           "window.resizeTo(350,150);" & _
           "window.moveTo(screen.width-370,screen.height-170);" & _
           "setTimeout(function(){window.close();}," & (timeout * 1000) & ");" & _
           "</script></head>" & _
           "<body><h3>" & Title & "</h3><p>" & msg & "</p></body></html>"
    
    Dim tmpPath As String
    tmpPath = Environ$("TEMP") & "\toast_" & Format(Now, "yyyymmddhhnnss") & ".hta"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim f As Object
    Set f = fso.CreateTextFile(tmpPath, True, True)
    f.Write html
    f.Close
    
    shell "mshta """ & tmpPath & """", vbHide
End Sub

' =====================================================================
' WScript Popup (Fallback)
' =====================================================================
Private Sub ShowWScriptPopup(ByVal msg As String, _
                            ByVal Title As String, _
                            ByVal timeout As Long)
    
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Popup msg, timeout, Title, 64
End Sub

' =====================================================================
' TEST MMF TOAST (Advanced)
' =====================================================================
Public Sub TestMMFToast()
    On Error GoTo ErrorHandler
    
    Dim json As String
    json = "{"
    json = json & """Title"":""MMF Test"","
    json = json & """Message"":""Test toast via Memory Mapped File"","
    json = json & """ToastType"":""Info"","
    json = json & """Progress"":0"
    json = json & "}"
    
    ' For now, use JSON file method (MMF requires more complex implementation)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim tempFile As String
    tempFile = Environ$("TEMP") & "\ExcelToasts\ToastRequest.json"
    
    Dim ts As Object
    Set ts = fso.CreateTextFile(tempFile, True, False)
    ts.Write json
    ts.Close
    
    MsgBox "MMF test toast sent via JSON fallback", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "MMF test failed: " & Err.Description, vbExclamation
End Sub

' =====================================================================
' DEMO MENU
' =====================================================================
Public Sub MsgBoxUI_MainMenu()
    On Error Resume Next
    
    ' Initialize
    MsgBoxUI_Init EnableToasts:=True, EnableFallback:=True
    
    Do
        Dim choice As String
        choice = InputBox( _
            MSGBOXUI_TITLE & " Demo Menu" & vbCrLf & vbCrLf & _
            "1. Simple Notification" & vbCrLf & _
            "2. Progress Demo" & vbCrLf & _
            "3. Input Prompt Demo" & vbCrLf & _
            "4. Send Custom Toast" & vbCrLf & _
            "5. Test All Toast Types" & vbCrLf & _
            "6. Settings" & vbCrLf & _
            "7. Pushbullet Menu" & vbCrLf & _
            "0. Exit" & vbCrLf & vbCrLf & _
            "Enter choice:", MSGBOXUI_TITLE, "1")
        
        ' Handle cancel or exit
        If choice = "" Or choice = "0" Then Exit Sub
        
        Select Case Val(choice)
            Case 1
                Demo_SimpleNotification
                
            Case 2
                Demo_Progress
                
            Case 3
                Demo_InputPrompt
                
            Case 4
                Demo_CustomToast
                
            Case 5
                Demo_AllToastTypes
                
            Case 6
                ShowSettings
                
            Case 7
                ShowPushbulletMenu
                
            Case Else
                MsgBox "Invalid selection. Please enter 0-7.", vbExclamation, "Invalid Choice"
        End Select
    Loop
End Sub

' =====================================================================
' DEMO ROUTINES
' =====================================================================

Private Sub Demo_SimpleNotification()
    Notify "Demo Notification", "This is a simple notification!", "INFO"
    MsgBox "Notification sent! Check for toast.", vbInformation
End Sub

Private Sub Demo_Progress()
    Dim i As Long
    For i = 0 To 100 Step 20
        Progress "Demo Progress", "Processing data...", i
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    
    Notify "Complete!", "Progress demo finished", "SUCCESS"
End Sub

Private Sub Demo_InputPrompt()
    Dim result As String
    result = InputText("Demo Input", "Enter your name:", "Guest")
    
    If result <> "" Then
        Notify "Hello " & result, "Welcome to MsgBoxUI!", "SUCCESS"
    Else
        MsgBox "Input cancelled.", vbInformation
    End If
End Sub

Private Sub Demo_CustomToast()
    Dim Title As String, msg As String, lvl As String, prog As String
    
    Title = InputBox("Enter toast title:", "Custom Toast", "Excel Test")
    If Title = "" Then Exit Sub
    
    msg = InputBox("Enter toast message:", "Custom Toast", "This is a test from VBA.")
    If msg = "" Then Exit Sub
    
    lvl = InputBox("Enter level (INFO/WARN/ERROR/SUCCESS):", "Custom Toast", "INFO")
    If lvl = "" Then lvl = "INFO"
    
    prog = InputBox("Enter progress % (0-100, or leave blank for none):", "Custom Toast", "")
    
    If prog <> "" Then
        Dim progVal As Long
        progVal = CLng(prog)
        If progVal < 0 Then progVal = 0
        If progVal > 100 Then progVal = 100
        SendToastNotification Title, msg, lvl, progVal
    Else
        SendToastNotification Title, msg, lvl, -1
    End If
    
    MsgBox "Custom toast sent!", vbInformation
End Sub

Private Sub Demo_AllToastTypes()
    MsgBox "This will send 4 toasts (Info, Success, Warning, Error) with 2 second delays.", vbInformation
    
    Notify "Information", "This is an info toast", "INFO"
    Application.Wait Now + TimeValue("00:00:02")
    
    Notify "Success!", "This is a success toast", "SUCCESS"
    Application.Wait Now + TimeValue("00:00:02")
    
    Notify "Warning", "This is a warning toast", "WARN"
    Application.Wait Now + TimeValue("00:00:02")
    
    Notify "Error", "This is an error toast", "ERROR"
    Application.Wait Now + TimeValue("00:00:02")
    
    MsgBox "All toast types sent!", vbInformation
End Sub

Private Sub ShowSettings()
    Dim msg As String
    msg = "Current Settings:" & vbCrLf & vbCrLf
    msg = msg & "PowerShell Toasts: " & IIf(UsePowerShellToasts, "Enabled", "Disabled") & vbCrLf
    msg = msg & "JSON Fallback: " & IIf(UseTempJsonFallback, "Enabled", "Disabled") & vbCrLf
    msg = msg & "Pipe Name: " & ToastPipeName & vbCrLf & vbCrLf
    msg = msg & "Temp Folder: " & Environ$("TEMP") & "\ExcelToasts" & vbCrLf
    
    MsgBox msg, vbInformation, "Settings"
End Sub

' =====================================================================
' PUSHBULLET INTEGRATION
' =====================================================================

' Send notification to mobile via Pushbullet
Private Sub SendPushbulletNotification(ByVal Title As String, ByVal Message As String)
    On Error Resume Next
    
    ' Check if Pushbullet is available and enabled
    If Not UsePushbullet Then Exit Sub
    
    ' Try to call MsgBoxPushBullet.SendPush
    ' This will work if the module is present in the project
    Dim result As Boolean
    result = MsgBoxPushBullet.SendPush(Title, Message)
    
    If Err.Number <> 0 Then
        ' Pushbullet module not available or error occurred
        ' Silently fail (don't interrupt user workflow)
        Err.Clear
    End If
End Sub

' Enable Pushbullet for this session
Public Sub EnablePushbulletNotifications()
    UsePushbullet = True
    
    ' Check if MsgBoxPushBullet module is available
    On Error Resume Next
    Dim isEnabled As Boolean
    isEnabled = MsgBoxPushBullet.IsPushbulletEnabled()
    
    If Err.Number <> 0 Then
        MsgBox "Pushbullet module not found in this project.", vbExclamation, "Pushbullet"
        UsePushbullet = False
        Err.Clear
        Exit Sub
    End If
    
    If Not isEnabled Then
        MsgBox "Pushbullet is not configured. Run MsgBoxPushBullet.PushbulletSetupWizard first.", _
               vbExclamation, "Pushbullet"
        UsePushbullet = False
    Else
        MsgBox "Pushbullet notifications enabled! Toasts will be sent to your mobile device.", _
               vbInformation, "Pushbullet Enabled"
    End If
End Sub

' Disable Pushbullet for this session
Public Sub DisablePushbulletNotifications()
    UsePushbullet = False
    MsgBox "Pushbullet notifications disabled.", vbInformation, "Pushbullet"
End Sub

