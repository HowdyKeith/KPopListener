# KPopListener

A powerful Windows notification toolkit built with **VBA** and **PowerShell**.

KPopListener helps you create rich toast notifications, unified message boxes with multiple backends, progress toasts, clipboard monitoring, TTS, tray icons, and more — directly from Excel, Access, or PowerShell.

## Features

- Modern Windows toast notifications (with progress, queuing, templates, and analytics)
- Unified `MsgBox` system with many fallback methods (Toast, WinRT, MSHTA, Python, PushBullet, etc.)
- PowerShell modules for advanced toasts, clipboard saving, TTS, and tray icons
- VBA helper modules for easy integration into Office applications
- Setup, logging, credentials, and ribbon UI support

## Repository Structure

- **VBA modules** (`.bas` and `.cls` files) – Core logic for Excel/Access
- **PowerShell scripts** (`.ps1` / `.psm1`) – Standalone or called from VBA
- Key files: `Setup.bas`, `ToastMasterDemo.bas`, `ToastExample.bas`, `MsgBoxUnified.bas`, `ToastSender.bas`

## Getting Started

### For VBA (Excel / Access)

1. Open your workbook or database.
2. Press `Alt + F11` to open the Visual Basic Editor.
3. Import the modules you need (especially `Setup.bas`, `ToastSender.bas`, `ToastHelpers.bas`, `MsgBoxUnified.bas`, and any class modules like `clsToastNotification.cls`).
4. Run `Setup` once to initialize registry settings and required components.

### For PowerShell

```powershell
Import-Module .\KPopupToastAdvanced.psm1
Detailed VBA Examples
1. Basic Toast Notification (using ToastSender)
vbaSub ShowSimpleToast()
    Call SendToast("KPopListener", "Hello from VBA!", "This is a basic toast notification.")
End Sub
2. Toast with Longer Message and Duration (from ToastExample.bas style)
vbaSub ShowDetailedToast()
    Dim Title As String
    Dim Message As String
    
    Title = "KPopListener Update"
    Message = "A new notification has been processed successfully." & vbCrLf & _
              "Time: " & Now & vbCrLf & _
              "Status: Complete"
    
    ' Send toast with custom duration (Short / Long)
    Call SendToast(Title, Message, "", "Long")
End Sub
3. Advanced Toast using ToastMasterDemo style
vbaSub ShowAdvancedToast()
    ' Example similar to ToastMasterDemo.bas
    Dim Toast As Object  ' or use the class if imported: Dim Toast As clsToastNotification
    
    ' Simple call using helpers
    Call ShowToastWithTemplate("Success", _
                               "Operation completed successfully!", _
                               "You can now continue with the next step.", _
                               "Info")   ' or "Success", "Warning", "Error"
End Sub
4. Unified Message Box (Recommended – works across many backends)
vbaSub TestMsgBoxUnified()
    Dim Response As VbMsgBoxResult
    
    Response = MsgBoxUnified( _
        Prompt:="Do you want to continue with the operation?", _
        Buttons:=vbYesNo + vbQuestion, _
        Title:="KPopListener Confirmation", _
        Backend:="Toast" )   ' Options: "Toast", "WinRT", "MSHTA", "Python", "PushBullet", etc.
    
    If Response = vbYes Then
        MsgBoxUnified "You chose Yes!", vbInformation, "Result"
    Else
        MsgBoxUnified "You chose No.", vbExclamation, "Result"
    End If
End Sub
5. Progress Toast (using clsToastProgress)
vbaSub ShowProgressToast()
    Dim ProgressToast As clsToastProgress
    
    Set ProgressToast = New clsToastProgress
    ProgressToast.Title = "Processing Files"
    ProgressToast.Message = "Please wait..."
    ProgressToast.Show
    
    ' Simulate work
    Dim i As Integer
    For i = 1 To 100
        ProgressToast.UpdateProgress i, "Step " & i & " of 100"
        Application.Wait (Now + TimeValue("0:00:01"))   ' 1 second delay for demo
    Next i
    
    ProgressToast.Complete "All files processed successfully!"
End Sub
6. Using Setup and Logging
vbaSub InitializeKPopListener()
    Call Setup   ' Runs initial configuration (registry, etc.)
    Call LogEvent("KPopListener initialized successfully", "Info")
    
    MsgBoxUnified "Setup complete! You can now use toasts and unified messages.", vbInformation
End Sub
7. Toast with Analytics / Queue Management
vbaSub SendToastWithAnalytics()
    Dim Analytics As clsToastAnalytics
    Set Analytics = New clsToastAnalytics
    
    Call SendToast("Analytics Test", "This toast will be tracked.", "Long")
    Analytics.TrackToast "ToastID123", "User clicked the action button"
End Sub
Tip: Start by importing and running ToastMasterDemo.bas and ToastExample.bas — they contain ready-to-run demonstration code.
PowerShell Examples
PowerShell# Advanced toast
Show-KPopupToast -Title "KPopListener" -Message "Hello from PowerShell!" -Duration Long

# Clipboard saver (monitor and save clipboard changes)
.\KPopupClipSaverAdvanced.psm1
Contributing
Contributions are welcome! Feel free to improve examples, add new backends, fix bugs, or enhance documentation.
License
Open source (add your preferred license if desired).

Made for notification lovers and Office automation enthusiasts.
