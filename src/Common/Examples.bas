Attribute VB_Name = "Examples"
'@Folder("Common")
Option Explicit

Private Token As String
Private RefreshToken As String

Public Sub ExampleGetToken()

    ' Example how to get token and refresh token for Microsoft Graph Api
    ' Application have to be registered in Azure Portal, please see:
    ' https://docs.microsoft.com/en-us/graph/auth-register-app-v2
    '
    ' Uses authorization code flow. More information:
    ' https://docs.microsoft.com/en-us/graph/auth-v2-user
    '
    
    On Error GoTo ErrHandler
    Dim Self As String
    Self = "ExampleGetToken"
    
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
    
    With New IeTokenProvider
        .Init "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize?client_id={client_id}&response_type=code&redirect_uri={redirect_uri}&response_mode=query&scope={scope}&state={state}", _
              "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token", _
              "client_id={client_id}&client_secret={client_secret}&scope={scope}&code={code}&redirect_uri={redirect_uri}&grant_type={grant_type}", _
              LoginTimeout:=60, settings:=settings
        .GetToken
        Token = .Token
        RefreshToken = .RefreshToken
    End With
    
    Debug.Print "Token", Token
    Debug.Print "RefreshToken", RefreshToken
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error!" & vbCrLf & vbCrLf & "Error description: " & Err.Description & vbCrLf & "Error source: " & Err.Source, vbExclamation, "Error!"

End Sub

Public Sub ExampleTokenPrompt()

    ' Prompts for token
    
    On Error GoTo ErrHandler
    Dim Self As String
    Self = "ExampleTokenPrompt"
    
    Dim tokenForm As TokenUserForm
    Set tokenForm = New TokenUserForm
    With tokenForm
        .Show
        If .OK Then
            Token = .TokenTextBox.value
        Else
            Token = vbNullString
        End If
    End With

    Exit Sub
    
ErrHandler:
    MsgBox "Error!" & vbCrLf & vbCrLf & "Error description: " & Err.Description & vbCrLf & "Error source: " & Err.Source, vbExclamation, "Error!"

End Sub

Public Sub ExampleOneDriveExplorer()

    ' Example how to use VBA OneDrive Explorer

    On Error GoTo ErrHandler
    Dim Self As String
    Self = "ExampleOneDriveExplorer"
    
    ' Choose one
    ExampleTokenPrompt
'    ExampleGetToken

    If Len(Token) = 0 Then Exit Sub
    
    Dim entryPointPath As String
    entryPointPath = "https://graph.microsoft.com/v1.0/me/drive/root/"

    Dim explorer As OneDriveFileExplorer
    Set explorer = New OneDriveFileExplorer
    explorer.Display entryPointPath:=entryPointPath, Token:=Token, userFormTitle:="Select file", allowMultiselect:=False, selectMode:=ESelectModeFilesOnly

    If Not explorer.IsCancelled Then
        ' Printing selected items' id, path
        If Not explorer.SelectedItems Is Nothing Then
            If explorer.SelectedItems.Count <> 0 Then
                Dim item As IDriveItem
                For Each item In explorer.SelectedItems
                    Debug.Print item.Id, item.path
                Next item
                
                Dim odFile As OneDriveFile
                Set odFile = explorer.SelectedItems(1)
                ExampleDownloadFile odFile
            End If
        End If
    End If

    Exit Sub
    
ErrHandler:
    Select Case Err.Number
    Case ErrorCodes.Unauthorized
        ' invalid or expired token or insufficent permissions
        
    Case Else
        MsgBox "Error!" & vbCrLf & vbCrLf & "Error description: " & Err.Description & vbCrLf & "Error source: " & Err.Source, vbExclamation, "Error!"
        
    End Select

End Sub

Public Sub ExampleOneDriveExplorerSharedWithMe()

    ' Example how to use VBA OneDrive Explorer

    On Error GoTo ErrHandler
    Dim Self As String
    Self = "ExampleOneDriveExplorer"
    
    ' Choose one
    ExampleTokenPrompt
'    ExampleGetToken
    
    If Len(Token) = 0 Then Exit Sub
    
    Dim entryPointPath As String
    entryPointPath = "https://graph.microsoft.com/v1.0/me/drive/SharedWithMe/"

    Dim explorer As OneDriveFileExplorer
    Set explorer = New OneDriveFileExplorer
    explorer.Display entryPointPath:=entryPointPath, Token:=Token, userFormTitle:="Select file", allowMultiselect:=False, selectMode:=ESelectModeFilesOnly

    If Not explorer.IsCancelled Then
        ' Printing selected items' id, path
        If Not explorer.SelectedItems Is Nothing Then
            Dim item As IDriveItem
            For Each item In explorer.SelectedItems
                Debug.Print item.Id, item.path
            Next item
            
            Dim odFile As OneDriveFile
            Set odFile = explorer.SelectedItems(1)
            ExampleDownloadFile odFile
        End If
    End If

    Exit Sub
    
ErrHandler:
    Select Case Err.Number
    Case ErrorCodes.Unauthorized
        ' invalid or expired token or insufficent permissions
        
    Case Else
        MsgBox "Error!" & vbCrLf & vbCrLf & "Error description: " & Err.Description & vbCrLf & "Error source: " & Err.Source, vbExclamation, "Error!"
        
    End Select

End Sub

Public Sub ExampleDownloadFile(ByRef file As OneDriveFile)

    ' Example how to download file from OneDrive

    On Error GoTo ErrHandler
    Dim Self As String
    Self = "ExampleDownloadFile"
    
    Dim request As WinHttp.WinHttpRequest
    Set request = New WinHttp.WinHttpRequest
    request.Open "GET", file.DownloadUrl, False
    request.Send
    
    Dim path As String
    path = ThisWorkbook.path & "\" & file.Name
    FileIO.WriteBinaryFile path, request.ResponseBody
    
    MsgBox "File downloaded to: " & path, vbInformation

    Exit Sub
    
ErrHandler:
    MsgBox "Error!" & vbCrLf & vbCrLf & "Error description: " & Err.Description & vbCrLf & "Error source: " & Err.Source, vbExclamation, "Error!"
    
End Sub

Private Sub WriteBinaryFile(ByVal path As String, ByVal varByteArray As Variant)

    On Error GoTo ErrHandler
    
    Dim stream As ADODB.stream
    Set stream = New ADODB.stream
    stream.Type = adTypeBinary
    stream.Open
    stream.Write varByteArray
    stream.SaveToFile path, adSaveCreateNotExist
    stream.Close
    
    Exit Sub
    
ErrHandler:
    stream.Close
    Err.Raise Err.Number, Err.Source & "FileIO.WriteBinaryFile", Err.Description

End Sub
