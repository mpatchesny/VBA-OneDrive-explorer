VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MicrosoftGraphApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Service")
Option Explicit

Implements IApi

Private token As String

Public Property Get Self() As MicrosoftGraphApi
    Set Self = Me
End Property
Private Property Get IApi_Self() As IApi
    Set IApi_Self = Self
End Property

Public Sub Init(ByVal cToken As String)
    GuardClauses.IsEmptyString cToken, "Token"
    token = cToken
End Sub

Public Function GetItem(ByVal id As String) As String
    ' ...
End Function
Private Function IApi_GetItem(ByVal id As String) As String
    IApi_GetItem = GetItem(id)
End Function

Public Function GetItems(ByVal parentId As String) As String
    ' ...
End Function
Private Function IApi_GetItems(ByVal parentId As String) As String
    IApi_GetItems = GetItems(parentId)
End Function

Private Function GetRequest(ByVal cToken As String, ByVal cUrl As String) As WinHttpRequest

    Dim request As WinHttp.WinHttpRequest
    Set request = New WinHttp.WinHttpRequest
    request.Option(4) = 13056 ' ignore all errors
    request.Open "GET", cUrl, False
    request.setRequestHeader "Authorization", "Bearer " & cToken
    Set GetRequest = request

End Function

