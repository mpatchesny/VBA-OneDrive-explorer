VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneDriveFileExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Service")
Option Explicit

Private cSelectedItems As Collection
Private cIsCancelled As Boolean

Public Property Get SelectedItems() As Collection
    Set SelectedItems = cSelectedItems
End Property

Public Property Get IsCancelled() As Boolean
    IsCancelled = cIsCancelled
End Property

Public Property Get Self() As OneDriveFileExplorer
    Set Self = Me
End Property

Public Sub Display(ByRef entryPointPath As String, _
                   ByVal Token As String, _
                   ByVal userFormTitle As String, _
                   ByVal allowMultiselect As Boolean, _
                   ByVal selectMode As ESelectMode)
    
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".Display"
    
    GuardClauses.IsEmptyString entryPointPath, "Entry point path"
    GuardClauses.IsEmptyString Token, "Token"
    GuardClauses.IsEmptyString userFormTitle, "User form title"
    
    Dim api As IApi
    With New MicrosoftGraphApi
        .Init Token
        Set api = .Self
    End With
    
    Dim provider As IItemProvider
    With New OneDriveItemProvider
        .Init New OneDriveFileFactory, New OneDriveFolderFactory, api
        Set provider = .Self
    End With
    
    Dim entryPoint  As IExplorerViewModel
    Set entryPoint = GetIExplorerViewModel(entryPointPath, provider)

    Dim controller As IExplorerController
    With New ExplorerControllerFactory
        Set controller = .NewExplorerController(entryPoint, userFormTitle, allowMultiselect, selectMode)
    End With
    
    controller.Display
    cIsCancelled = controller.IsCancelled
    If Not IsCancelled Then Set cSelectedItems = controller.SelectedItems
    
    Exit Sub
    
ErrHandler:
    Dim msg As String
    msg = "Error: " & Err.Description & " (" & Err.Source & ")"
    MsgBox msg, vbExclamation, "Error"
    
End Sub

Private Function GetIExplorerViewModel(ByVal entryPointPath As String, ByRef provider As IItemProvider) As IExplorerViewModel

    Dim parent As IDriveItem
    With New OneDriveFolder
        .Init "0", "0", "Root folder", Nothing, 0, entryPointPath, Now, provider
        Set parent = .Self
    End With

    Dim items As Collection
    Set items = GetItemsSafe(provider, parent)
    
    If items Is Nothing Then
        Dim item As IDriveItem
        Set item = provider.GetItemByPath(entryPointPath)
        Dim fld As IFolder
        Set fld = item
        Set items = fld.GetChildren
    End If
    
    Dim entryPoint As IExplorerViewModel
    Set entryPoint = New ExplorerViewModel
    entryPoint.SetItems items
    Set GetIExplorerViewModel = entryPoint

End Function

Private Function GetItemsSafe(ByRef provider As IItemProvider, ByRef parent As IDriveItem) As Collection
    
    On Error GoTo ErrHandler
    Dim items As Collection
    Set items = provider.GetItems(parent)
    Exit Function
    
ErrHandler:
    Err.Clear
    
End Function
