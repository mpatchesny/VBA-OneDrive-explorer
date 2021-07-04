VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExplorerView 
   Caption         =   "UserForm1"
   ClientHeight    =   5820
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

Public Property Get Self() As ExplorerView
    Set Self = Me
End Property
Private Property Get IExplorerView_Self() As IExplorerView
    Set IExplorerView_Self = Self
End Property

Public Sub Display()
    Me.Display
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

Private Sub ListView_DblClick()
    ' TODO
End Sub

Private Sub UserForm_Resize()
    ' TODO
End Sub
