VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExplorerController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Controllers")
Option Explicit

Implements IExplorerController

Private Type TFields
    View As IExplorerView
    Model As IExplorerViewModel
    IsDisplayed As Boolean
    IsCancelled As Boolean
    SelectedItems As Collection
End Type
Private this As TFields

Public Property Get IsDisplayed() As Boolean
    IsDisplayed = this.IsDisplayed
End Property
Private Property Get IExplorerController_IsDisplayed() As Boolean
    IExplorerController_IsDisplayed = IsDisplayed
End Property
Private Property Let IsDisplayed(ByVal value As Boolean)
    this.IsDisplayed = value
End Property

Public Property Get IsCancelled() As Boolean
    IsCancelled = this.IsCancelled
End Property
Private Property Get IExplorerController_IsCancelled() As Boolean
    IExplorerController_IsCancelled = IsCancelled
End Property
Private Property Let IsCancelled(ByVal value As Boolean)
    this.IsCancelled = value
End Property

Public Property Get SelectedItems() As Collection
    Set SelectedItems = this.Model.SelectedItems
End Property
Private Property Get IExplorerController_SelectedItems() As Collection
    Set IExplorerController_SelectedItems = SelectedItems
End Property
Private Property Let SelectedItems(ByRef value As Collection)
    Set this.SelectedItems = value
End Property

Public Property Get Self() As ExplorerController
    Set Self = Me
End Property
Private Property Get IExplorerController_Self() As IExplorerController
    Set IExplorerController_Self = Self
End Property

Public Sub Init(ByRef cView As IExplorerView, ByRef cModel As Object)
    Set this.View = cView
    Set this.Model = cModel
End Sub

Public Sub Display()
    
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".Self"
    
    IsDisplayed = True
    this.View.Display
    IsCancelled = this.View.IsCancelled
    IsDisplayed = False
    Exit Sub
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description
    
End Sub
Private Sub IExplorerController_Display()
    Display
End Sub





