VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITokenProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Service.Interfaces")
'@Interface
Option Explicit

Public Property Get Token() As String
End Property

Public Property Get RefreshToken() As String
End Property

Public Property Get Self() As ITokenProvider
End Property

Public Sub GetToken()
End Sub
