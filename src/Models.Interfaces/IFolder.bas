VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Models.Interfaces")
'@Interface
Option Explicit

Public Property Get Self() As IFolder
End Property

Public Property Get id() As String
End Property

Public Property Get name() As String
End Property

Public Property Get ChildrenCount() As Long
End Property

Public Function GetChildren() As Collection
End Function

Public Property Get path() As String
End Property

