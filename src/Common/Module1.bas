Attribute VB_Name = "Module1"
'@Folder("Common")
Option Explicit

Sub Test()

    On Error GoTo ErrHandler
    Dim Self As String
    Self = ".test"
    
    Dim entryPoint As IDriveItem
    Dim col As Collection
    Dim col2 As Collection
    
    With New FileSystemItemProvider
        Set entryPoint = .GetItem("C:\Users\strielok\Downloads\")
        Set col = .GetItems(entryPoint)
    End With
    
    DebugPrintItemCol col
    
    Dim item As IDriveItem
    Dim newFolder As IFolder
    For Each item In col
        If item.IsFolder Then
            Set newFolder = item
            Set col2 = newFolder.GetChildren
            DebugPrintItemCol col2
        End If
    Next item
    
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
