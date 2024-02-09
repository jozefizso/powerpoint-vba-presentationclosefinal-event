Option Explicit

Public WithEvents App As Application

Sub App_PresentationClose(ByVal Pres As Presentation)
    MsgBox "App_PresentationClose event.", vbInformation, "AppClass macro"
End Sub

' This event is not fired on PowerPoint on macOS
Sub App_PresentationCloseFinal(ByVal Pres As Presentation)
    MsgBox "App_PresentationCloseFinal event.", vbInformation, "AppClass macro"
End Sub

Sub App_PresentationBeforeClose(ByVal Pres As Presentation, ByRef Cancel As Boolean)
    MsgBox "App_PresentationBeforeClose event.", vbInformation, "AppClass macro"
End Sub
