VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExplorerView 
   Caption         =   "UserForm1"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "ExplorerView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExplorerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Views")
Implements IExplorerView

Private Type TFields
    IsCancelled As Boolean
End Type
Private this As TFields

Public Property Get IsCancelled() As Boolean
    IsCancelled = False
End Property
Private Property Get IExplorerView_IsCancelled() As Boolean
    IExplorerView_IsCancelled = IsCancelled
End Property
Private Property Let IsCancelled(ByVal value As Boolean)
    this.IsCancelled = value
End Property

Public Property Get Self() As ExplorerView
    Set Self = Me
End Property
Private Property Get IExplorerView_Self() As IExplorerView
    Set IExplorerView_Self = Self
End Property

Public Sub Init(ByVal title As String)
    Me.Caption = title
End Sub

Public Sub Display()
    Me.RefreshButton.Caption = ChrW(&HE72C)
    Me.Show
End Sub
Private Sub IExplorerView_Display()
    Display
End Sub

Public Sub HideView()
    Me.Hide
End Sub
Private Sub IExplorerView_HideView()
    HideView
End Sub

Private Sub OkButton_Click()
    OK
End Sub

Private Sub RefreshButton_Click()
    ' TODO
End Sub

Private Sub UserForm_QueryClose(cCancel As Integer, CloseMode As Integer)
    If Not IsCancelled Then
        cCancel = True
        Cancel
    End If
End Sub

Private Sub ListView_DblClick()
    ' TODO
End Sub

Private Sub UserForm_Resize()
    ' TODO
End Sub

Public Sub OK()
    IsCancelled = False
    HideView
End Sub

Public Sub Cancel()
    IsCancelled = True
    HideView
End Sub
