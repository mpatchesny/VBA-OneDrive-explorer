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

Public Function NewExplorerController(ByRef entryPoint As IExplorerViewModel, _
                                      ByVal userFormTitle As String, _
                                      ByVal multiselect As Boolean, _
                                      Optional ByVal selectMode As ESelectMode = ESelectMode.ESelectModeAll) As IExplorerController
                                      
    GuardClauses.IsNothing entryPoint, "Entry point"
    GuardClauses.IsEmptyString userFormTitle, "User form title"
    
    Dim View As IExplorerView
    With New ExplorerView
        .Init entryPoint, userFormTitle, multiselect, selectMode
        Set View = .Self
    End With
    
    With New ExplorerController
        .Init View, entryPoint
        Set NewExplorerController = .Self
    End With
                                        
End Function

Private Function IExplorerControllerFactory_NewExplorerController(ByRef entryPoint As IDriveItem, _
                                                                 ByVal userFormTitle As String, _
                                                                 ByVal multiselect As Boolean, _
                                                                 Optional ByVal selectMode As ESelectMode = ESelectMode.ESelectModeAll) As IExplorerController
    Set IExplorerControllerFactory_NewExplorerController = NewExplorerController(entryPoint, userFormTitle, multiselect, selectMode)
End Function

