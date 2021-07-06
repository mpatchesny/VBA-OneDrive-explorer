VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IExplorerControllerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Factory.Interfaces")
'@Interface
Option Explicit

Public Property Get Self() As IExplorerControllerFactory
End Property

Public Function NewExplorerController(ByRef entryPoint As IDriveItem, _
                                      ByVal userformTitle As String, _
                                      ByVal multiselect As Boolean, _
                                      Optional ByVal selectMode As ESelectMode = ESelectMode.ESelectModeAll) As IExplorerController
End Function
