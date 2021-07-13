Attribute VB_Name = "Module1"
'@Folder("Common")
Option Explicit

Sub Test()

    On Error GoTo ErrHandler
    Dim Self As String
    Self = ".test"
    
    Dim token As String
    token = FileIO.ReadFileAlt(ThisWorkbook.Path & "\..\token.txt", "UTF-8")
    
    Dim explorer As OneDriveFileExplorer
    Set explorer = New OneDriveFileExplorer
    explorer.Display "C:\Users\strielok\Desktop", token, "Select file", True, ESelectModeAll
    
    If Not explorer.IsCancelled Then
        Dim SelectedItems As Collection
        Set SelectedItems = explorer.SelectedItems
        DebugPrintItemCol SelectedItems
    End If
    
    Exit Sub
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Sub

Private Function DebugPrintItemCol(ByRef col As Collection)

    If Not col Is Nothing Then
        Dim item As IDriveItem
        For Each item In col
            Debug.Print item.Path
        Next item
    End If
    
End Function
