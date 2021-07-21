VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Models.Interfaces")
'@Interface
Option Explicit

Public Property Get Self() As IFile
End Property

Public Property Get Id() As String
End Property

Public Property Get DriveId() As String
End Property

Public Property Get Name() As String
End Property

Public Property Get Size() As String
End Property

Public Property Get CreatedTime() As Date
End Property

Public Property Get LastModifiedTime() As Date
End Property

Public Property Get path() As String
End Property
