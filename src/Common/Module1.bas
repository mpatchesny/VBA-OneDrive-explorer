Attribute VB_Name = "Module1"
'@Folder("Common")
Option Explicit

Public Sub Start()

    On Error GoTo ErrHandler

    Dim token As String
    token = "" ' paste your token here

    Dim explorer As OneDriveFileExplorer
    Set explorer = New OneDriveFileExplorer
    explorer.Display entryPointPath:="https://graph.microsoft.com/v1.0/me/drive/root/", token:=token, userformTitle:="Select file", allowMultiselect:=True, selectMode:=ESelectModeAll

    If Not explorer.IsCancelled Then
        Dim SelectedItems As Collection
        Set SelectedItems = explorer.SelectedItems
    End If

    Exit Sub
    
ErrHandler:
    MsgBox "Error!" & vbCrLf & vbCrLf & "Error description: " & err.Description & vbCrLf & "Error source: " & err.Source, vbExclamation, "Error!"

End Sub

Private Sub DebugPrintSelectedItems(ByRef col As Collection)

    If Not col Is Nothing Then
        Dim item As IDriveItem
        For Each item In col
            Debug.Print item.Id, item.Path
        Next item
    End If
    
End Sub

