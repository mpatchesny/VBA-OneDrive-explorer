Attribute VB_Name = "Module2"
'@Folder("Common")
Option Explicit

Public Sub ApiTest()

    Dim token As String
    token = FileIO.ReadFileAlt(ThisWorkbook.path & "\..\token.txt", "UTF-8")
    
    Dim url As String
    url = "https://graph.microsoft.com/v1.0/me/"
    
    Dim request As WinHttp.WinHttpRequest
    Set request = New WinHttp.WinHttpRequest
    request.Option(4) = 13056 ' ignore all errors
    request.Open "GET", url, False
    request.setRequestHeader "Authorization", "Bearer " & token
    request.send ("")
    
    Debug.Print request.Status
    
    If request.Status = 200 Then
        Dim response As String
        response = request.ResponseText
    End If

End Sub
