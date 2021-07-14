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

Public Sub Display(ByVal entryPointPath As String, _
                   ByVal token As String, _
                   ByVal userformTitle As String, _
                   ByVal allowMultiselect As Boolean, _
                   ByVal selectMode As ESelectMode)
    
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".Display"
    
    GuardClauses.IsEmptyString entryPointPath, "Entry point"
    GuardClauses.IsEmptyString token, "Token"
    GuardClauses.IsEmptyString userformTitle, "User form title"
    
    Dim api As IApi
    With New MicrosoftGraphApi
        .Init token
        Set api = .Self
    End With
    
    Dim provider As IItemProvider
    With New OneDriveItemProvider
        .Init New OneDriveFileFactory, New OneDriveFolderFactory, api
        Set provider = .Self
    End With
    
    Dim entryPoint As IDriveItem
    Set entryPoint = provider.GetItemByPath(entryPointPath)

    Dim controller As IExplorerController
    With New ExplorerControllerFactory
        Set controller = .NewExplorerController(entryPoint, userformTitle, allowMultiselect, selectMode)
    End With
    
    controller.Display
    cIsCancelled = controller.IsCancelled
    If Not IsCancelled Then Set cSelectedItems = controller.SelectedItems
    
    Exit Sub
    
ErrHandler:
    Dim msg As String
    msg = "Error: " & err.Description & " (" & err.Source & ")"
    MsgBox msg, vbExclamation, "Error"
    
End Sub
