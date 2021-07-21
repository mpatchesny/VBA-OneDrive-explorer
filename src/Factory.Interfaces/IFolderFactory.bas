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

Public Function newFolder(ByVal Id As String, _
                            ByVal DriveId As String, _
                            ByVal Name As String, _
                            ByRef Parent As IDriveItem, _
                            ByVal ChildrenCount As Long, _
                            ByVal path As String, _
                            ByVal LastModifiedTime As Date, _
                            ByRef provider As IItemProvider) As IFolder
End Function

Public Function NewFolderFromDictionary(ByRef d As Scripting.IDictionary, _
                                        ByRef Parent As IDriveItem, _
                                        ByRef provider As IItemProvider) As IFolder
End Function






