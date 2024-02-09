Option Explicit

Dim oApp As New AppClass

Public Sub Auto_Open()
    ' Connect the Application object to our AppClass
    Set oApp.App = Application
End Sub
