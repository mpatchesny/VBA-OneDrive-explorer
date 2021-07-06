VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneDriveFolderFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Factory")
Option Explicit

Implements IFolderFactory

Public Property Get Self() As OneDriveFolderFactory
    Set Self = Me
End Property
Private Property Get IFolderFactory_Self() As IFolderFactory
    Set IFolderFactory_Self = Self
End Property

Public Function NewFolder(ByVal Id As String, _
                        ByVal Name As String, _
                        ByRef Parent As IDriveItem, _
                        ByVal ChildrenCount As Long, _
                        ByVal path As String, _
                        ByRef provider As IItemProvider) As IFolder
    With New OneDriveFolder
        .Init Id, Name, Parent, ChildrenCount, path, provider
        Set NewFolder = .Self
    End With
End Function
Private Function IFolderFactory_NewFolder(ByVal Id As String, _
                                        ByVal Name As String, _
                                        ByRef Parent As IDriveItem, _
                                        ByVal ChildrenCount As Long, _
                                        ByVal path As String, _
                                        ByRef provider As IItemProvider) As IFolder
    Set IFolderFactory_NewFolder = NewFolder(Id, Name, Parent, ChildrenCount, path, provider)
End Function

Public Function NewFolderFromJsonString(ByVal json As String) As IFolder
    ' TODO
End Function
Private Function IFolderFactory_NewFolderFromJsonString(ByVal json As String) As IFolder
    Set IFolderFactory_NewFolderFromJsonString = NewFolderFromJsonString(json)
End Function





