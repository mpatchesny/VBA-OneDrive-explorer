VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExplorerViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Models")
Option Explicit

Implements IExplorerViewModel

Private Type TFields
    items As Collection
    CurrentItem As IDriveItem
    SelectedItems As Collection
End Type
Private this As TFields

Public Property Get Self() As ExplorerViewModel
    Set Self = Me
End Property
Private Property Get IExplorerViewModel_Self() As IExplorerViewModel
    Set IExplorerViewModel_Self = Self
End Property

Public Property Get items() As Collection
    Set items = this.items
End Property
Private Property Get IExplorerViewModel_Items() As Collection
    Set IExplorerViewModel_Items = items
End Property
Private Property Let items(ByRef value As Collection)
    Set this.items = value
End Property

Public Property Get CurrentItem() As IDriveItem
    Set CurrentItem = this.CurrentItem
End Property
Private Property Get IExplorerViewModel_CurrentItem() As IDriveItem
    Set IExplorerViewModel_CurrentItem = CurrentItem
End Property
Private Property Let CurrentItem(ByRef value As IDriveItem)
    Set this.CurrentItem = value
End Property

Public Property Get SelectedItems() As Collection
    Set SelectedItems = this.SelectedItems
End Property
Private Property Get IExplorerViewModel_SelectedItems() As Collection
    Set IExplorerViewModel_SelectedItems = SelectedItems
End Property
Private Property Let SelectedItems(ByRef value As Collection)
    Set this.SelectedItems = value
End Property

Public Sub Init(ByRef cItems As Collection, ByRef cCurrentItem As IDriveItem, ByRef cSelectedItems As Collection)
    SetItems cItems
    SetCurrentItem cCurrentItem
    SetSelectedItems cSelectedItems
End Sub

Public Sub SetItems(ByRef value As Collection)
    items = value
End Sub
Private Sub IExplorerViewModel_SetItems(value As Collection)
    SetItems value
End Sub

Public Sub SetCurrentItem(ByRef value As IDriveItem)
    CurrentItem = value
End Sub
Private Sub IExplorerViewModel_SetCurrentItem(value As IDriveItem)
    SetCurrentItem value
End Sub

Public Sub SetSelectedItems(ByRef value As Collection)
    SelectedItems = value
End Sub
Private Sub IExplorerViewModel_SetSelectedItems(value As Collection)
    SetSelectedItems value
End Sub

