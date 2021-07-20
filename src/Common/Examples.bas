Attribute VB_Name = "Examples"
'@Folder("Common")
Option Explicit

Public Sub ExampleGetToken()

    ' Example how to get token and refresh token for Microsoft Graph Api
    ' Application have to be registered in Azure Portal, please see:
    ' https://docs.microsoft.com/en-us/graph/auth-register-app-v2
    '
    ' Uses authorization code flow. More information:
    ' https://docs.microsoft.com/en-us/graph/auth-v2-user
    '
    
    On Error GoTo ErrHandler
    
    Dim settings As RequestSettings
    With New RequestSettings
        .ClientId = "client_id_from_microsoft_azure_portal"
        .ClientSecret = "client_secret_from_microsoft_azure_portal"
        .GrantType = "authorization_code"
        .RedirectUri = "http%3A%2F%2Flocalhost%2F"
        .ResponseMode = "query"
        .ResponseType = "code"
        .Scope = "offline_access%20user.read%20Files.ReadWrite.All"
        .State = "12345"
        .Tenant = "consumers" ' consumers/ organizations/ common
        Set settings = .Self
    End With
    
    Dim Token As String
    Dim RefreshToken As String
    With New IeTokenProvider
        .Init "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize?client_id={client_id}&response_type=code&redirect_uri={redirect_uri}&response_mode=query&scope={scope}&state={state}", _
              "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token", _
              "client_id={client_id}&client_secret={client_secret}&scope={scope}&code={code}&redirect_uri={redirect_uri}&grant_type={grant_type}", _
            LoginTimeout:=100, settings:=settings
        .GetToken
        Token = .Token
        RefreshToken = .RefreshToken
    End With
    
    Debug.Print "Token", Token
    Debug.Print "RefreshToken", RefreshToken
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error!" & vbCrLf & vbCrLf & "Error description: " & err.Description & vbCrLf & "Error source: " & err.Source, vbExclamation, "Error!"

End Sub

Public Sub ExampleOneDriveExplorer()

    ' Example how to use VBA OneDrive Explorer

    On Error GoTo ErrHandler
    
    Dim Token As String
    Token = "" ' paste your token here
    
    Dim entryPointPath As String
    entryPointPath = "https://graph.microsoft.com/v1.0/me/drive/root"

    Dim explorer As OneDriveFileExplorer
    Set explorer = New OneDriveFileExplorer
    explorer.Display entryPointPath:=entryPointPath, Token:=Token, userformTitle:="Select file", allowMultiselect:=True, selectMode:=ESelectModeAll

    If Not explorer.IsCancelled Then
        ' Printing selected items' id, path
        If Not explorer.SelectedItems Is Nothing Then
            Dim item As IDriveItem
            For Each item In explorer.SelectedItems
                Debug.Print item.Id, item.Path
            Next item
        End If
    End If

    Exit Sub
    
ErrHandler:
    MsgBox "Error!" & vbCrLf & vbCrLf & "Error description: " & err.Description & vbCrLf & "Error source: " & err.Source, vbExclamation, "Error!"

End Sub
