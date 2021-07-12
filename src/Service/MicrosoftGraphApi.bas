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

Private cResponseStatus As Long
Private cResponse As String
Private token As String

Public Property Get ResponseStatus() As Long
    ResponseStatus = cResponseStatus
End Property
Private Property Get IApi_ResponseStatus() As Long
    IApi_ResponseStatus = ResponseStatus
End Property

Public Property Get Response() As String
    Response = cResponse
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

Public Function GetItem(ByVal id As String) As String
        
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItem"
    
    Dim query As String
    query = ""
    
    With GetRequest(token, query)
        .send ("")
        cResponseStatus = .Status
        cResponse = .ResponseText
        
        If cResponseStatus = 200 Then
            Dim Response As String
            Response = .ResponseText
            GetItem = Response
            
        Else
            RaiseBadResponseError cResponseStatus
            
        End If
    End With
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description
    

End Function
Private Function IApi_GetItem(ByVal id As String) As String
    IApi_GetItem = GetItem(id)
End Function

Public Function GetItems(ByVal parentId As String) As String
        
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItems"
    
    Dim query As String
    query = ""
    
    With GetRequest(token, query)
        .send ("")
        cResponseStatus = .Status
        cResponse = .ResponseText
        
        If cResponseStatus = 200 Then
            Dim Response As String
            Response = .ResponseText
            GetItems = Response
            
        Else
            RaiseBadResponseError cResponseStatus
            
        End If
    End With
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description
    
End Function
Private Function IApi_GetItems(ByVal parentId As String) As String
    IApi_GetItems = GetItems(parentId)
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

Private Sub RaiseBadResponseError(ByVal resStatus As Long)

    Select Case resStatus
    Case 400
        err.Raise ErrorCodes.BadRequest, Self, "Bad request"
        
    Case 401
        err.Raise ErrorCodes.Unauthorized, Self, "Unauthorized"
        
    Case 403
        err.Raise ErrorCodes.Forbidden, Self, "Forbidden"
        
    Case 404
        err.Raise ErrorCodes.NotFound, Self, "Not found"
        
    Case 405
        err.Raise ErrorCodes.MethodNotAllowed, Self, "Method not allowed"
        
    Case 406
        err.Raise ErrorCodes.NotAcceptable, Self, "Not acceptable"
        
    Case 412
        err.Raise ErrorCodes.PreconditionFailed, Self, "Precondition failed"
    
    Case 500
        err.Raise ErrorCodes.InternalServerError, Self, "Internal server error"
    
    Case Else
        err.Raise ErrorCodes.BadResponse, Self, "Bad response status " & resStatus
        
    End Select

End Sub

