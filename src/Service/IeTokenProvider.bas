VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IeTokenProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Service")
Option Explicit

Implements ITokenProvider

Private Type TFields
    UrlGetTemplate As String
    UrlPostTemplate As String
    PostBodyTemplate As String
    LoginTimeout As Long
    settings As RequestSettings
    Token As String
    RefreshToken As String
End Type
Private this As TFields

Public Property Get Token() As String
    Token = this.Token
End Property
Private Property Get ITokenProvider_Token() As String
    ITokenProvider_Token = Token
End Property

Public Property Get RefreshToken() As String
    RefreshToken = this.RefreshToken
End Property
Private Property Get ITokenProvider_RefreshToken() As String
    ITokenProvider_RefreshToken = RefreshToken
End Property

Public Property Get Self() As IeTokenProvider
    Set Self = Me
End Property
Private Property Get ITokenProvider_Self() As ITokenProvider
    Set ITokenProvider_Self = Self
End Property

Public Sub Init(ByVal UrlGetTemplate As String, ByVal UrlPostTemplate As String, ByVal PostBodyTemplate As String, ByVal LoginTimeout As Long, ByVal settings As RequestSettings)
    
    GuardClauses.IsEmptyString UrlGetTemplate, "Url GET template"
    GuardClauses.IsEmptyString UrlPostTemplate, "Url POST template"
    GuardClauses.IsEmptyString PostBodyTemplate, "POST body template"
    GuardClauses.IsZero LoginTimeout, "Login timeout"
    GuardClauses.IsNothing settings, "Request settings"
    
    this.UrlGetTemplate = UrlGetTemplate
    this.UrlPostTemplate = UrlPostTemplate
    this.PostBodyTemplate = PostBodyTemplate
    this.LoginTimeout = LoginTimeout
    Set this.settings = settings
    
End Sub

Public Sub GetToken()

    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetToken"
    
    Dim url As String
    url = this.UrlGetTemplate
    url = Replace(url, "{tenant}", this.settings.Tenant)
    url = Replace(url, "{client_id}", this.settings.ClientId)
    url = Replace(url, "{response_type}", this.settings.ResponseType)
    url = Replace(url, "{redirect_uri}", this.settings.RedirectUri)
    url = Replace(url, "{response_mode}", this.settings.ResponseMode)
    url = Replace(url, "{scope}", this.settings.Scope)
    url = Replace(url, "{state}", this.settings.State)

    Dim responseCode As String
    responseCode = GetCode(url, this.LoginTimeout)
    
    Dim requestUrl As String
    requestUrl = this.UrlPostTemplate
    requestUrl = Replace(requestUrl, "{tenant}", this.settings.Tenant)
    
    Dim tokens As Variant
    tokens = GetTokens(requestUrl, this.PostBodyTemplate, responseCode, this.settings)
    this.Token = tokens(0)
    this.RefreshToken = tokens(1)
    
    Exit Sub
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description
    
End Sub
Private Sub ITokenProvider_GetToken()
    GetToken
End Sub

Private Function GetCode(ByVal url As String, ByVal timeout As Long) As String

    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetCode"
    
    With New InternetExplorer
        .Visible = True
        .navigate url
    End With
    
    Application.Wait (DateAdd("s", 3, Now))

    Dim elapsedTimeInSeconds As Long
    Dim responseUrl As Variant
    Dim urls As Variant
    
    Dim match As Boolean
    Do While Not match
        urls = GetInternetExplorerOpenedURLs
        If Not IsEmpty(urls) Then
            For Each responseUrl In urls
                match = (responseUrl Like "http://localhost/?code=*")
                If match Then Exit For
            Next responseUrl
            
        Else
            err.Raise ErrorCodes.InternetExplorerIsClosed, Self, "Internet Explorer is closed"
            
        End If
        
        If Not match Then
            Application.Wait (DateAdd("s", 1, Now))
            elapsedTimeInSeconds = elapsedTimeInSeconds + 1
            If elapsedTimeInSeconds > timeout Then err.Raise ErrorCodes.TimeoutExceeded, Self, "Failed to login in " & timeout & " seconds"
        End If
    Loop
    
    TerminateIEProcess
    
    GetCode = GetCodeFromUrl(responseUrl)
    Exit Function
    
ErrHandler:
    TerminateIEProcess
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Function

Private Function GetTokens(ByVal requestUrl As String, ByVal bodyTemplate As String, ByVal responseCode As String, ByRef settings As RequestSettings) As Variant

    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetTokens"
    
    Dim Token As String
    Dim RefreshToken As String
    
    Dim body As String
    body = bodyTemplate
    body = Replace(body, "{client_id}", settings.ClientId)
    body = Replace(body, "{client_secret}", settings.ClientSecret)
    body = Replace(body, "{scope}", settings.Scope)
    body = Replace(body, "{code}", responseCode)
    body = Replace(body, "{redirect_uri}", settings.RedirectUri)
    body = Replace(body, "{grant_type}", settings.GrantType)
    
    With New WinHttp.WinHttpRequest
        .Option(4) = 13056 ' ignore all errors
        .Open "POST", requestUrl, False
        .SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
        .Send body
        
        If .Status = 200 Then
            Dim json As String
            json = .responseText
            
            Dim dict As Scripting.Dictionary
            If Utils.TryParseJson(json, dict) Then
                Token = dict("access_token")
                RefreshToken = dict("refresh_token")
                GetTokens = Array(Token, RefreshToken)
                
            Else
                err.Raise ErrorCodes.BadResponse, Self, "Tokens not found in response text"
                
            End If
            
        Else
            RaiseBadResponseError .Status, .responseText, Self
            
        End If
    End With
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Function

' Helpers
Private Sub TerminateIEProcess()

    On Error Resume Next
    Dim service As Object
    Set service = GetObject("winmgmts:\\.\root\cimv2")

    Dim processes As Object
    Set processes = service.ExecQuery("Select * From Win32_Process")
    
    Dim process As Object
    For Each process In processes
        If process.Name = "iexplore.exe" Then process.Terminate
    Next

End Sub

Private Function GetInternetExplorerOpenedURLs() As Variant
    
    Dim shell As Object
    Set shell = CreateObject("Shell.Application")
    
    Dim urls As Variant
    
    Dim wnd As Object
    For Each wnd In shell.Windows
        If InStr(1, wnd, "Internet Explorer", vbTextCompare) <> 0 Then
            urls = ArrayAddItem(urls, wnd.LocationURL)
        End If
    Next
    
    GetInternetExplorerOpenedURLs = urls
    
End Function

Private Function GetCodeFromUrl(ByVal url As String) As String

    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetCodeFromUrl"

    Dim match As IMatchCollection
    Dim submatch As IMatch2
    
    With New RegExp
        .Pattern = "code=(.+)&"
        If .Test(url) Then
            Set match = .Execute(url)
            If match.Count = 1 Then
                GetCodeFromUrl = match.item(0).SubMatches(0)
                
            Else
                err.Raise ErrorCodes.FailToGetCodeFromUrl, Self, "Failed to get code from url: " & url
                
            End If
            
        Else
            err.Raise ErrorCodes.FailToGetCodeFromUrl, Self, "Failed to get code from url: " & url
            
        End If
    End With
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Function

Private Function ArrayAddItem(ByVal arr As Variant, ByVal item As Variant) As Variant

    Dim lngMax As Long
    If IsEmpty(arr) Then
        ReDim arr(0)
        lngMax = 0
    Else
        lngMax = UBound(arr, 1) + 1
        ReDim Preserve arr(lngMax)
    End If
    
    arr(lngMax) = item
    ArrayAddItem = arr

End Function

Private Sub RaiseBadResponseError(ByVal resStatus As Long, ByVal responseText As String, ByVal methodName As String)

    Select Case resStatus
    Case 400
        err.Raise ErrorCodes.BadRequest, methodName, "Bad request (" & responseText & ")"
        
    Case 401
        err.Raise ErrorCodes.Unauthorized, methodName, "Unauthorized (" & responseText & ")"
        
    Case 403
        err.Raise ErrorCodes.Forbidden, methodName, "Forbidden (" & responseText & ")"
        
    Case 404
        err.Raise ErrorCodes.NotFound, methodName, "Not found (" & responseText & ")"
        
    Case 405
        err.Raise ErrorCodes.MethodNotAllowed, methodName, "Method not allowed (" & responseText & ")"
        
    Case 406
        err.Raise ErrorCodes.NotAcceptable, methodName, "Not acceptable (" & responseText & ")"
        
    Case 412
        err.Raise ErrorCodes.PreconditionFailed, methodName, "Precondition failed (" & responseText & ")"
    
    Case 500
        err.Raise ErrorCodes.InternalServerError, methodName, "Internal server error (" & responseText & ")"
    
    Case Else
        err.Raise ErrorCodes.BadResponse, methodName, "Bad response status " & resStatus & " (" & responseText & ")"
        
    End Select

End Sub
