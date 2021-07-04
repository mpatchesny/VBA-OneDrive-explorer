Attribute VB_Name = "Module1"
'@Folder("Common")
Option Explicit

Sub Test()

    On Error GoTo ErrHandler
    Dim Self As String
    Self = ".test"
    
    Dim entryPoint As IDriveItem
    
    Exit Sub
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Sub
