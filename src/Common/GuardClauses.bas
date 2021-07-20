Attribute VB_Name = "GuardClauses"
'@Folder("Common")
Option Explicit

Public Sub IsNothing(ByRef o As Object, ByVal Name As String)
    If o Is Nothing Then err.Raise ErrorCodes.ArgumentNotNull, "GuardClauses.IsNothing", Name & " cannot be nothing"
End Sub

Public Sub IsEmptyString(ByVal value As String, ByVal Name As String)
    If Len(value) = 0 Then
        err.Raise ErrorCodes.ArgumentNotNull, "GuardClauses.IsEmptyString", Name & " cannot be empty string"
    End If
End Sub

Public Sub IsZero(ByVal value As Long, ByVal Name As String)
    If Len(value) = 0 Then
        err.Raise ErrorCodes.ArgumentNotNull, "GuardClauses.IsZero", Name & " cannot be zero"
    End If
End Sub
