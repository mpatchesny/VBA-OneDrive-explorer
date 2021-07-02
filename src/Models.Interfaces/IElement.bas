VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Models.Interfaces")
'@Interface
Option Explicit

Public Enum ElementType
    folder = 1
    File = 2
End Enum

Public Property Get Self() As IElement
End Property

Public Property Get Parent() As IElement
End Property

Public Property Get Id() As String
End Property

Public Property Get ElementType() As ElementType
End Property

Public Property Get DisplayName() As String
End Property

Public Function GetChildrenCount() As Long
End Function

Public Function GetChildren() As Collection
End Function

Public Sub Download(ByVal saveToPath As String)
End Sub

