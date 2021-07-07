Attribute VB_Name = "GuardClauses"
'@Folder("Common")
Option Explicit

Public Sub IsNothing(ByRef o As Object, ByVal name As String)
    If o Is Nothing Then err.Raise ErrorCodes.ArgumentNotNull, "GuardClauses.IsNothing", name & " cannot be nothing"
End Sub

Public Sub IsEmptyString(ByVal value As String, ByVal name As String)
    If Len(value) = 0 Then
        err.Raise ErrorCodes.ArgumentNotNull, "GuardClauses.IsNullOrEmpty", name & " cannot be empty string"
    End If
End Sub
