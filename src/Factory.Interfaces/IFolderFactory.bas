VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFolderFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Factory.Interfaces")
'@Interface
Option Explicit

Public Property Get Self() As IFolderFactory
End Property

Public Function NewFolder(ByVal id As String, _
                        ByVal name As String, _
                        ByRef Parent As IDriveItem, _
                        ByVal ChildrenCount As Long, _
                        ByVal path As String, _
                        ByRef provider As IItemProvider) As IFolder
End Function

Public Function NewFolderFromJsonString(ByVal json As String) As IFolder
End Function






