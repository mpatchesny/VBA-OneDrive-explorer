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

Private Type TResponse
    ResponseStatus As Long
    ResponseText As String
End Type

Private thisResponse As TResponse
Private token As String

Public Property Get ResponseStatus() As Long
    ResponseStatus = thisResponse.ResponseStatus
End Property
Private Property Get IApi_ResponseStatus() As Long
    IApi_ResponseStatus = ResponseStatus
End Property

Public Property Get Response() As String
    Response = thisResponse.ResponseText
End Property
Private Property Get IApi_Response() As String
    IApi_Response = Response
End Property

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

Public Function GetItem(ByVal Id As String, ByVal DriveId As String) As String
        
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItem"
    
    Dim query As String
    query = GetQuery(Id, DriveId)
    
    Dim req As WinHttpRequest
    Set req = GetRequest(token, query)
    
    thisResponse = ExecuteRequest(req)
    With thisResponse
        If ResponseStatus = 200 Then
            GetItem = .ResponseText
            
        Else
            ' TODO: log bad response status
            RaiseBadResponseError ResponseStatus, "GetItems"
            
        End If
    End With
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Function
Private Function IApi_GetItem(ByVal Id As String, ByVal DriveId As String) As String
    IApi_GetItem = GetItem(Id, DriveId)
End Function

Public Function GetItems(ByVal parentId As String, ByVal DriveId As String) As String
        
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItems"
    
    Dim query As String
    query = GetQueryChildren(parentId, DriveId)
    
    Dim req As WinHttpRequest
    Set req = GetRequest(token, query)
    
    thisResponse = ExecuteRequest(req)
    With thisResponse
        If ResponseStatus = 200 Then
            GetItems = .ResponseText
            
        Else
            ' TODO: log bad response status
            RaiseBadResponseError ResponseStatus, "GetItems"
            
        End If
    End With
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description
    
End Function
Private Function IApi_GetItems(ByVal parentId As String, ByVal DriveId As String) As String
    IApi_GetItems = GetItems(parentId, DriveId)
End Function

Public Function ExecuteQuery(ByVal query As String) As String
        
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItems"
    
    Dim req As WinHttpRequest
    Set req = GetRequest(token, query)
    
    thisResponse = ExecuteRequest(req)
    With thisResponse
        If ResponseStatus = 200 Then
            ExecuteQuery = .ResponseText
            
        Else
            ' TODO: log bad response status
            RaiseBadResponseError ResponseStatus, "GetItems"
            
        End If
    End With
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description
    
End Function
Private Function IApi_ExecuteQuery(ByVal query As String) As String
    IApi_ExecuteQuery = ExecuteQuery(query)
End Function

Private Function GetQuery(ByVal Id As String, ByVal DriveId As String) As String
    ' FIXME: tym sie powinna zajmowaæ jakaœ osobna klasa, query provider albo coœ takiego
    ' nie mo¿na zak³¹daæ, ze zawsze bêdziemy znajdowaæ siê w swoim onedrive, a nie np. w
    ' elementach ktore s¹ nam udostêpnione
    Dim query As String
    query = "https://graph.microsoft.com/v1.0/me/drives/{DriveId}/items/{Id}"
    query = Replace(query, "{DriveId}", DriveId)
    query = Replace(query, "{Id}", Id)
    GetQuery = query
End Function

Private Function GetQueryChildren(ByVal Id As String, ByVal DriveId As String) As String
    Dim query As String
    query = GetQuery(Id, DriveId)
    query = query & "/children"
    GetQueryChildren = query
End Function

Private Function GetRequest(ByVal cToken As String, ByVal cUrl As String) As WinHttpRequest
    Dim request As WinHttp.WinHttpRequest
    Set request = New WinHttp.WinHttpRequest
    With request
        .Option(4) = 13056 ' ignore all errors
        .Open "GET", cUrl, False
        .setRequestHeader "Authorization", "Bearer " & cToken
    End With
    Set GetRequest = request
End Function

Private Function ExecuteRequest(ByRef req As WinHttpRequest) As TResponse

    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".ExecuteRequest"

    With req
        .Send ("")
        
        Dim t As TResponse
        t.ResponseStatus = .Status
        t.ResponseText = .ResponseText
        ExecuteRequest = t
    End With
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description
    
End Function

Private Sub RaiseBadResponseError(ByVal resStatus As Long, ByVal methodName As String)

    Select Case resStatus
    Case 400
        err.Raise ErrorCodes.BadRequest, methodName, "Bad request"
        
    Case 401
        err.Raise ErrorCodes.Unauthorized, methodName, "Unauthorized"
        
    Case 403
        err.Raise ErrorCodes.Forbidden, methodName, "Forbidden"
        
    Case 404
        err.Raise ErrorCodes.NotFound, methodName, "Not found"
        
    Case 405
        err.Raise ErrorCodes.MethodNotAllowed, methodName, "Method not allowed"
        
    Case 406
        err.Raise ErrorCodes.NotAcceptable, methodName, "Not acceptable"
        
    Case 412
        err.Raise ErrorCodes.PreconditionFailed, methodName, "Precondition failed"
    
    Case 500
        err.Raise ErrorCodes.InternalServerError, methodName, "Internal server error"
    
    Case Else
        err.Raise ErrorCodes.BadResponse, methodName, "Bad response status " & resStatus
        
    End Select

End Sub

