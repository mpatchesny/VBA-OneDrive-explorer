VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TokenUserForm 
   Caption         =   "Paste token here"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8640
   OleObjectBlob   =   "TokenUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TokenUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Views")
Option Explicit

Private cOK As Boolean
Public Property Get OK() As Boolean
    OK = cOK
End Property

Private Sub OkButton_Click()
    If Len(TokenTextBox.text) = 0 Then
        MsgBox "Token cannot be empty.", vbExclamation
        Exit Sub
    End If
    cOK = True
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    cOK = False
End Sub

Private Sub UserForm_QueryClose(cCancel As Integer, CloseMode As Integer)
    cOK = False
    cCancel = True
    Me.Hide
End Sub
