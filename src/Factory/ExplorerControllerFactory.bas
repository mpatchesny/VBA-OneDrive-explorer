VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExplorerControllerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Factory")
Option Explicit

Implements IExplorerControllerFactory

Public Property Get Self() As ExplorerControllerFactory
    Set Self = Me
End Property
Private Property Get IExplorerControllerFactory_Self() As IExplorerControllerFactory
    Set IExplorerControllerFactory_Self = Self
End Property

Public Function NewExplorerController(ByRef entryPoint As IDriveItem, _
                                      ByVal userformTitle As String, _
                                      ByVal multiselect As Boolean, _
                                      Optional ByVal selectMode As ESelectMode = ESelectMode.ESelectModeAll) As IExplorerController
                                      
    GuardClauses.IsNothing entryPoint, TypeName(Me)
    GuardClauses.IsEmptyString userformTitle, "User form title"
                                        
    Dim Model As IExplorerViewModel
    With New ExplorerViewModel
        .Init Nothing, entryPoint, Nothing
        Set Model = .Self
    End With
    
    Dim View As IExplorerView
    With New ExplorerView
        .Init Model, userformTitle, multiselect, selectMode
        Set View = .Self
    End With
    
    With New ExplorerController
        .Init View, Model
        Set NewExplorerController = .Self
    End With
                                        
End Function

Private Function IExplorerControllerFactory_NewExplorerController(ByRef entryPoint As IDriveItem, _
                                                                 ByVal userformTitle As String, _
                                                                 ByVal multiselect As Boolean, _
                                                                 Optional ByVal selectMode As ESelectMode = ESelectMode.ESelectModeAll) As IExplorerController
    Set IExplorerControllerFactory_NewExplorerController = NewExplorerController(entryPoint, userformTitle, multiselect, selectMode)
End Function

