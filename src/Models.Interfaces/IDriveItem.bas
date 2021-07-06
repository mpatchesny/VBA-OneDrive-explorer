VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDriveItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Models.Interfaces")
'@Interface
Option Explicit

Public Property Get Self() As IDriveItem
End Property

Public Property Get Parent() As IDriveItem
End Property

Public Property Get Id() As String
End Property

Public Property Get IsFile() As Boolean
End Property

Public Property Get IsFolder() As Boolean
End Property

Public Property Get path() As String
End Property
