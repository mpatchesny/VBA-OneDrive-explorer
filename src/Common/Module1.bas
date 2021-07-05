Attribute VB_Name = "Module1"
'@Folder("Common")
Option Explicit

Sub Test()

    On Error GoTo ErrHandler
    Dim Self As String
    Self = ".test"
    
    Dim entryPoint As IDriveItem
    With New FileSystemItemProvider
        Set entryPoint = .GetItem("C:\Users\strielok\Downloads\")
    End With
    
    ' FIXME: factory
    Dim Model As IExplorerViewModel
    With New ExplorerViewModel
        .Init Nothing, entryPoint, Nothing
        Set Model = .Self
    End With
    
    Dim View As IExplorerView
    With New ExplorerView
        .Init Model, "Select file", False
        Set View = .Self
    End With
    
    Dim controller As IExplorerController
    With New ExplorerController
        .Init View, Model
        Set controller = .Self
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
