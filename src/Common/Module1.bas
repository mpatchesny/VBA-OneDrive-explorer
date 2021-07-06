Attribute VB_Name = "Module1"
'@Folder("Common")
Option Explicit

Sub Test()

    On Error GoTo ErrHandler
    Dim Self As String
    Self = ".test"
    
    Dim provider As IItemProvider
    With New FileSystemItemProvider
        .Init New OneDriveFileFactory, New OneDriveFolderFactory
        Set provider = .Self
    End With
    
    Dim entryPoint As IDriveItem
    Set entryPoint = provider.GetItem("C:\Users\strielok\Desktop")

    Dim controller As IExplorerController
    With New ExplorerControllerFactory
        Set controller = .NewExplorerController(entryPoint, "Select file", False)
    End With
    controller.Display
    
    Dim SelectedItems As Collection
    If Not controller.IsCancelled Then
        Set SelectedItems = controller.SelectedItems
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
            Debug.Print item.path
        Next item
    End If
    
End Function
