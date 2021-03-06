VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IItemProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Service.Interfaces")
'@Interface
Option Explicit

Public Property Get Self() As IItemProvider
End Property

Public Function GetItemById(ByVal Id As String, Optional ByRef Parent As IDriveItem) As IDriveItem
End Function

Public Function GetItemByPath(ByVal path As String, Optional ByRef Parent As IDriveItem) As IDriveItem
End Function

Public Function GetItems(ByRef Parent As IDriveItem) As Collection
End Function
