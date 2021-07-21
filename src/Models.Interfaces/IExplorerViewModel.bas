VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IExplorerViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Models.Interfaces")
'@Interface
Option Explicit

Public Property Get Self() As IExplorerViewModel
End Property

Public Property Get items() As Collection
End Property

Public Property Get currentItem() As IDriveItem
End Property

Public Property Get SelectedItems() As Collection
End Property

Public Sub SetItems(ByRef value As Collection)
End Sub

Public Sub SetCurrentItem(ByRef value As IDriveItem)
End Sub

Public Sub SetSelectedItems(ByRef value As Collection)
End Sub
