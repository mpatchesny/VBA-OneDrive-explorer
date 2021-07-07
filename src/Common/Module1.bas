Attribute VB_Name = "Module1"
'@Folder("Common")
Option Explicit

Sub Test()

    On Error GoTo ErrHandler
    Dim Self As String
    Self = ".test"
    
    Dim SelectedItems As Collection
    With New OneDriveFileExplorer
        .Display "C:\Users\strielok\Desktop", "x", "Select file", True, ESelectModeAll
        If Not .IsCancelled Then
            Set SelectedItems = .SelectedItems
            DebugPrintItemCol SelectedItems
        End If
    End With
    
    Exit Sub
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Sub

Private Function DebugPrintItemCol(ByRef col As Collection)

    If Not col Is Nothing Then
        Dim item As IDriveItem
        For Each item In col
            Debug.Print item.path
        Next item
    End If
    
End Function
