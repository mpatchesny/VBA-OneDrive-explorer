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
    responseText As String
End Type

Private thisResponse As TResponse
Private Token As String

Public Property Get ResponseStatus() As Long
    ResponseStatus = thisResponse.ResponseStatus
End Property
Private Property Get IApi_ResponseStatus() As Long
    IApi_ResponseStatus = ResponseStatus
End Property

Public Property Get response() As String
    response = thisResponse.responseText
End Property
Private Property Get IApi_Response() As String
    IApi_Response = response
End Property

Public Property Get Self() As MicrosoftGraphApi
    Set Self = Me
End Property
Private Property Get IApi_Self() As IApi
    Set IApi_Self = Self
End Property

Public Sub Init(ByVal cToken As String)
    GuardClauses.IsEmptyString cToken, "Token"
    Token = cToken
End Sub

Public Function GetItemById(ByVal Id As String, ByVal DriveId As String) As String
        
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItemById"
    
    Dim query As String
    query = GetQuery(Id, DriveId)
    ExecuteQuery (query)
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source & ";" & Self, Err.Description

End Function
Private Function IApi_GetItemById(ByVal Id As String, ByVal DriveId As String) As String
    IApi_GetItemById = GetItemById(Id, DriveId)
End Function

Public Function GetItemByPath(ByVal path As String) As String
        
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItemByPath"
    
    Dim query As String
    query = path
    ExecuteQuery (query)
    GetItemByPath = thisResponse.responseText
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source & ";" & Self, Err.Description

End Function
Private Function IApi_GetItemByPath(ByVal path As String) As String
    IApi_GetItemByPath = GetItemByPath(path)
End Function

Public Function GetItems(ByRef parent As IDriveItem) As String
        
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItems"
    
    Dim query As String
    If LCase(parent.path) Like "*/drive/sharedwithme*" Then
        query = parent.path
        
    Else
        query = GetQueryChildren(parent.Id, parent.DriveId)
        
    End If
    
    ExecuteQuery (query)
    GetItems = thisResponse.responseText
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source & ";" & Self, Err.Description
    
End Function
Private Function IApi_GetItems(ByRef parent As IDriveItem) As String
    IApi_GetItems = GetItems(parent)
End Function

Private Function ExecuteQuery(ByVal query As String) As String
        
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".ExecuteQuery"
    
    Dim req As WinHttpRequest
    Set req = GetRequest(Token, query)
    
    thisResponse = ExecuteRequest(req)
    With thisResponse
        If ResponseStatus = 200 Then
            ExecuteQuery = .responseText
            
        Else
            ' TODO: log bad response status
            RaiseBadResponseError ResponseStatus, .responseText, "GetItems"
            
        End If
    End With
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source & ";" & Self, Err.Description
    
End Function

Private Function GetRequest(ByVal cToken As String, ByVal cUrl As String) As WinHttpRequest
    Dim request As WinHttp.WinHttpRequest
    Set request = New WinHttp.WinHttpRequest
    With request
        .Option(4) = 13056 ' ignore all errors
        .Open "GET", cUrl, False
        .SetRequestHeader "Authorization", "Bearer " & cToken
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
        t.responseText = .responseText
        ExecuteRequest = t
    End With
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source & ";" & Self, Err.Description
    
End Function

' Helpers
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

Private Sub RaiseBadResponseError(ByVal resStatus As Long, ByVal responseText As String, ByVal methodName As String)

    Select Case resStatus
    Case 400
        Err.Raise ErrorCodes.BadRequest, methodName, "Bad request (" & responseText & ")"
        
    Case 401
        Err.Raise ErrorCodes.Unauthorized, methodName, "Unauthorized (" & responseText & ")"
        
    Case 403
        Err.Raise ErrorCodes.Forbidden, methodName, "Forbidden (" & responseText & ")"
        
    Case 404
        Err.Raise ErrorCodes.NotFound, methodName, "Not found (" & responseText & ")"
        
    Case 405
        Err.Raise ErrorCodes.MethodNotAllowed, methodName, "Method not allowed (" & responseText & ")"
        
    Case 406
        Err.Raise ErrorCodes.NotAcceptable, methodName, "Not acceptable (" & responseText & ")"
        
    Case 412
        Err.Raise ErrorCodes.PreconditionFailed, methodName, "Precondition failed (" & responseText & ")"
    
    Case 500
        Err.Raise ErrorCodes.InternalServerError, methodName, "Internal server error (" & responseText & ")"
    
    Case Else
        Err.Raise ErrorCodes.BadResponse, methodName, "Bad response status " & resStatus & " (" & responseText & ")"
        
    End Select

End Sub
