VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Service.Interfaces")
'@Interface
Option Explicit

Public Property Get ResponseStatus() As Long
End Property

Public Property Get response() As String
End Property

Public Property Get Self() As IApi
End Property

Public Function GetItem(ByVal Id As String, ByVal DriveId As String) As String
End Function

Public Function GetItems(ByVal parentId As String, ByVal DriveId As String) As String
End Function

Public Function ExecuteQuery(ByVal query As String) As String
End Function
