Attribute VB_Name = "ToastExample"
' ToastExample.bas
' Version: 3.3
' Purpose: Sample Toast Notifications for VBA
' Dependencies: Logs.bas, MsgBoxMSHTA.bas, MsgBoxUniversal.bas, Setup.bas, MsgBoxUI.bas
' Notes:
' - Consolidated test routines for ribbon & menu integration
' - Supports simple, image, advanced, MsgBoxMSHTA, and PowerShell progress toasts
' - Safe JSON handling via EscapeJson
' - Ribbon includes listener start/stop buttons
' - Uses Setup.TEMP_FOLDER for temp paths
' - Added ShowSimpleToast, ShowToastWithImage, ShowAdvancedToast, ShowToastPowerShell
' - Fixed syntax error in ProgressToastDemo
' - Fixed EscapeJson quote escaping
' Date: October 24, 2025
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'=========================
' LISTENER CONTROL SYSTEM
'=========================
Private ListenerStates As Object
Public RibbonUI As Object

Private Sub UpdateRibbonListenerLabels()
    If Not RibbonUI Is Nothing Then
        RibbonUI.InvalidateControl "btnTogglePS"
        RibbonUI.InvalidateControl "btnTogglePy"
    End If
End Sub

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set RibbonUI = ribbon
End Sub

Public Function GetListenerImage(control As IRibbonControl) As StdPicture
    InitListenerStates
    Dim isActive As Boolean, imgPath As String, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Select Case control.ID
        Case "btnTogglePS": isActive = ListenerStates("PowerShell")
        Case "btnTogglePy": isActive = ListenerStates("Python")
    End Select
    
    If isActive Then
        imgPath = Setup.TEMP_FOLDER & "\ToastRibbon_Active.ico"
    Else
        imgPath = Setup.TEMP_FOLDER & "\ToastRibbon_Inactive.ico"
    End If
    
    If Not fso.FileExists(imgPath) Then SaveTempListenerIcons
    Set GetListenerImage = LoadPicture(imgPath)
End Function

Private Sub SaveTempListenerIcons()
    Dim activeIcon As String, inactiveIcon As String, f As Integer
    Dim activeData As String, inactiveData As String
    
    activeIcon = Setup.TEMP_FOLDER & "\ToastRibbon_Active.ico"
    inactiveIcon = Setup.TEMP_FOLDER & "\ToastRibbon_Inactive.ico"
    
    ' Small embedded ICOs (placeholder base64)
    activeData = "AAABAAEAICAAAAEAIACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAQAEAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=="
    inactiveData = activeData
    
    f = FreeFile
    Open activeIcon For Binary As #f: Put #f, , Base64Decode(activeData): Close #f
    f = FreeFile
    Open inactiveIcon For Binary As #f: Put #f, , Base64Decode(inactiveData): Close #f
End Sub

Private Function Base64Decode(ByVal base64String As String) As Byte()
    Dim xml As Object, node As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = base64String
    Base64Decode = node.nodeTypedValue
End Function

Private Sub InitListenerStates()
    If ListenerStates Is Nothing Then
        Set ListenerStates = CreateObject("Scripting.Dictionary")
        ListenerStates("PowerShell") = False
        ListenerStates("Python") = False
    End If
End Sub

Private Sub ToggleListenerState(ByVal ListenerType As String)
    InitListenerStates
    Dim CurrentState As Boolean: CurrentState = ListenerStates(ListenerType)
    
    If ListenerType = "PowerShell" Then
        If CurrentState Then
            MsgBoxUniversal.StopToastListener
            Logs.DebugLog "[ToastExample] PowerShell listener stopped", "INFO"
            MsgBox "PowerShell listener stopped.", vbInformation
        Else
            MsgBoxUniversal.StartToastListener
            Logs.DebugLog "[ToastExample] PowerShell listener started", "INFO"
            MsgBox "PowerShell listener started.", vbInformation
        End If
    ElseIf ListenerType = "Python" Then
        If CurrentState Then
            MsgBoxUniversal.StopPythonListener
            Logs.DebugLog "[ToastExample] Python listener stopped", "INFO"
            MsgBox "Python listener stopped.", vbInformation
        Else
            MsgBoxUniversal.StartPythonListener
            Logs.DebugLog "[ToastExample] Python listener started", "INFO"
            MsgBox "Python listener started.", vbInformation
        End If
    End If
    
    ListenerStates(ListenerType) = Not CurrentState
    UpdateRibbonListenerLabels
End Sub

Public Sub StartPowerShellListener()
    ToggleListenerState "PowerShell"
End Sub

Public Sub StopPowerShellListener()
    If Not ListenerStates Is Nothing Then If ListenerStates("PowerShell") Then ToggleListenerState "PowerShell"
End Sub

Public Sub StartPythonListener()
    ToggleListenerState "Python"
End Sub

Public Sub StopPythonListener()
    If Not ListenerStates Is Nothing Then If ListenerStates("Python") Then ToggleListenerState "Python"
End Sub

'=========================
' Ribbon Installer
'=========================
Public Sub InstallToastRibbon()
    On Error GoTo ErrorHandler
    RemoveToastRibbon
    
    Dim ribbonXml As String
    ribbonXml = _
        "<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='OnRibbonLoad'>" & _
        "<ribbon><tabs>" & _
        "<tab id='tabToastDemo' label='Toast Demos'>" & _
        "<group id='grpToastSamples' label='Toast Samples'>" & _
        "<button id='btnSimpleToast' label='Simple Toast' onAction='Test_SimpleToast' size='large'/>" & _
        "<button id='btnToastWithImage' label='Toast with Image' onAction='Test_ToastWithImage' size='large'/>" & _
        "<button id='btnAdvancedToast' label='Advanced Toast' onAction='Test_AdvancedToast' size='large'/>" & _
        "<button id='btnMsgBoxMSHTA' label='MsgBoxMSHTA Demo' onAction='Test_MsgBoxMSHTA' size='large'/>" & _
        "<button id='btnProgressToastDemo' label='Progress Toast Demo' onAction='ProgressToastDemo' size='large'/>" & _
        "<button id='btnAllToasts' label='All Toasts' onAction='Test_AllToasts' size='large'/>" & _
        "<button id='btnTogglePS' label='Toggle PS Listener' onAction='StartPowerShellListener' size='large' getImage='GetListenerImage'/>" & _
        "<button id='btnTogglePy' label='Toggle Python Listener' onAction='StartPythonListener' size='large' getImage='GetListenerImage'/>" & _
        "</group></tab></tabs></ribbon></customUI>"
    
    ThisWorkbook.CustomXMLParts.Add ribbonXml
    MsgBox "Toast ribbon installed successfully!", vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "Error installing Toast ribbon: " & Err.Description, vbExclamation
End Sub

Public Sub RemoveToastRibbon()
    Dim xPart As CustomXMLPart
    For Each xPart In ThisWorkbook.CustomXMLParts
        If InStr(1, xPart.xml, "tabToastDemo") > 0 Then xPart.Delete
    Next xPart
End Sub

'=========================
' Toast Demos
'=========================
Public Sub ShowSimpleToast(ByVal Title As String, ByVal Message As String)
    On Error Resume Next
    MsgBoxUI.Notify Title, Message
    Logs.DebugLog "[ToastExample] Simple toast displayed: " & Title & " - " & Message, "INFO"
End Sub

Public Sub ShowToastWithImage(ByVal Title As String, ByVal Message As String, ByVal ImagePath As String)
    On Error Resume Next
    If IsImageFile(ImagePath) Then
        MsgBoxMSHTA.ShowToast Title, Message, "INFO", 4, "BR", "asterisk", ImagePath
    Else
        Logs.DebugLog "[ToastExample] Invalid image path: " & ImagePath, "WARN"
        MsgBoxUI.Notify Title, Message & vbCrLf & "(Image not shown: invalid path)"
    End If
    Logs.DebugLog "[ToastExample] Toast with image displayed: " & Title, "INFO"
End Sub

Public Sub ShowAdvancedToast(ByVal Title As String, ByVal Message As String, ByVal Details As String)
    On Error Resume Next
    MsgBoxUI.Notify Title, Message & vbCrLf & Details, "INFO", 5, "C"
    Logs.DebugLog "[ToastExample] Advanced toast displayed: " & Title, "INFO"
End Sub

Public Function ShowToastPowerShell( _
    ByVal Title As String, _
    ByVal Message As String, _
    ByVal durationSec As Long, _
    ByVal Level As String, _
    Optional ByVal ImagePath As String, _
    Optional ByVal Icon As String, _
    Optional ByVal Sound As String = "BEEP", _
    Optional ByVal HeroImage As String, _
    Optional ByVal Attribution As String, _
    Optional ByVal Scenario As String, _
    Optional ByVal IsProgress As Boolean = False, _
    Optional ByVal Position As String = "C", _
    Optional ByVal ProgressValue As Long = 0) As Boolean
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim requestFile As String: requestFile = Setup.TEMP_FOLDER & "\ToastRequest.json"
    
    ' Build JSON payload
    Dim json As String
    json = "{" & _
           """Title"": """ & EscapeJson(Title) & """," & _
           """Message"": """ & EscapeJson(Message) & """," & _
           """Level"": """ & UCase(Level) & """," & _
           """DurationSec"": " & durationSec & "," & _
           """IsProgress"": " & IIf(IsProgress, "true", "false") & "," & _
           """Position"": """ & Position & """," & _
           """ProgressValue"": " & ProgressValue & _
           IIf(ImagePath <> "", ",""ImagePath"": """ & EscapeJson(ImagePath) & """", "") & _
           IIf(Sound <> "", ",""Sound"": """ & Sound & """", "") & _
           "}"
    
    ' Write JSON to file
    Dim ts As Object: Set ts = fso.CreateTextFile(requestFile, True, True)
    ts.Write json
    ts.Close
    
    ' Notify via MsgBoxUniversal
    ShowToastPowerShell = MsgBoxUniversal.ShowToast(Title, Message, Level, durationSec, Position, Sound, ImagePath, IsProgress, ProgressValue)
    
    If ShowToastPowerShell Then
        Logs.DebugLog "[ToastExample] PowerShell toast sent: " & Title & " - " & Message, "INFO"
    Else
        Logs.DebugLog "[ToastExample] Failed to send PowerShell toast: " & Title, "ERROR"
    End If
End Function

Public Sub Test_SimpleToast()
    ShowSimpleToast "Hello from VBA!", "This is a simple notification"
End Sub

Public Sub Test_ToastWithImage()
    Dim imgPath As String: imgPath = "C:\Windows\Web\Wallpaper\Windows\img0.jpg"
    ShowToastWithImage "VBA Notification", "This notification references an image!", imgPath
End Sub

Public Sub Test_AdvancedToast()
    ShowAdvancedToast "Process Complete", "Your data has been processed successfully", "Time: " & Now & vbCrLf & "Records: 1,234"
End Sub

Public Sub Test_AllToasts()
    Dim Timestamp As String: Timestamp = "[" & GetStardate & "]"
    Logs.DebugLog Timestamp & " [ToastSamples] Testing all notification types...", "INFO"
Test_SimpleToast:     Application.Wait Now + TimeValue("00:00:02")
Test_ToastWithImage:     Application.Wait Now + TimeValue("00:00:02")
Test_AdvancedToast:     Application.Wait Now + TimeValue("00:00:02")
Test_MsgBoxMSHTA:     Application.Wait Now + TimeValue("00:00:02")
    ProgressToastDemo
    Logs.DebugLog Timestamp & " [ToastSamples] All demos complete.", "INFO"
End Sub

'=========================
' Progress Toast Demo (PowerShell)
'=========================
Public Sub ProgressToastDemo()
    Dim i As Long, fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim requestFile As String: requestFile = Setup.TEMP_FOLDER & "\ToastRequest.json"
    
    If Not MsgBoxUniversal.PowershellListenerRunning() Then
        MsgBoxUniversal.StartToastListener
        Application.Wait Now + TimeValue("00:00:02")
    End If
    
    If ShowToastPowerShell("Processing Task", "Starting process...", 0, "INFO", , , "BEEP", , , , True, "C", 0) Then
        Debug.Print "Progress toast initiated"
    Else
        MsgBox "Failed to initiate progress toast - is listener running?", vbExclamation
        Exit Sub
    End If
    
    For i = 10 To 100 Step 10
        ShowToastPowerShell "Processing Task", "Processing: " & i & "%", 0, "INFO", , , "BEEP", , , , True, "C", i
        Dim maxWait As Long, j As Long
        maxWait = 30
        For j = 1 To maxWait
            If Not fso.FileExists(requestFile) Then Exit For
            Sleep 100: DoEvents
        Next j
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    
    ShowToastPowerShell "Task Complete", "Processing finished!", 3, "INFO", , , "BEEP", , , , , "C", 100
    MsgBox "Progress demo complete!", vbInformation
End Sub

'=========================
' MsgBoxMSHTA Demo
'=========================
Public Sub Test_MsgBoxMSHTA()
    Dim Timestamp As String: Timestamp = "[" & GetStardate & "]"
    Logs.DebugLog Timestamp & " [ToastSamples] MsgBoxMSHTA demo starting...", "INFO"
    
    MsgBoxMSHTA.ShowToast "Info Toast", "This is a normal info toast.", "INFO", 4, "BR", "asterisk"
    Application.Wait Now + TimeValue("00:00:02")
    MsgBoxMSHTA.ShowToast "Warning", "This is a warning toast.", "WARN", 4, "TR", "exclamation"
    Application.Wait Now + TimeValue("00:00:02")
    MsgBoxMSHTA.ShowToast "Error!", "This is an error toast.", "ERROR", 5, "TL", "hand"
    Application.Wait Now + TimeValue("00:00:02")
    MsgBoxMSHTA.ShowToast "Success", "This operation completed successfully.", "SUCCESS", 4, "BL", "asterisk"
    Application.Wait Now + TimeValue("00:00:02")
    
    Dim i As Long
    For i = 0 To 100 Step 20
        MsgBoxMSHTA.ShowProgressToast "Progress Demo", "Processing task...", i, "BR"
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    
    MsgBoxMSHTA.CleanupTempFiles
    Logs.DebugLog Timestamp & " [ToastSamples] MsgBoxMSHTA demo complete.", "INFO"
End Sub

'=========================
' Helper Functions
'=========================
Private Function EscapeJson(ByVal s As String) As String
    s = Replace(s, "\", "\\") ' Escape backslashes
    s = Replace(s, """", "\\""") ' Escape double quotes
    s = Replace(s, vbCrLf, "\n") ' Replace line breaks with \n
    EscapeJson = s
End Function

Private Function GetStardate() As String
    GetStardate = "2025." & Format(Now, "ddd.hh")
End Function

Private Function IsObjectNothing(ByVal obj As Object) As Boolean
    On Error Resume Next
    IsObjectNothing = (obj Is Nothing)
End Function

Private Function IsImageFile(ByVal filePath As String) As Boolean
    On Error Resume Next
    Dim ext As String
    ext = LCase(CreateObject("Scripting.FileSystemObject").GetExtensionName(filePath))
    IsImageFile = (ext = "png" Or ext = "jpg" Or ext = "jpeg" Or ext = "gif" Or ext = "bmp")
End Function



