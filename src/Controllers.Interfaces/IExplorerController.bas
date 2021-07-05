VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IExplorerController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Controllers.Interfaces")
'@Interface
Option Explicit

Public Property Get IsDisplayed() As Boolean
End Property

Public Property Get IsCancelled() As Boolean
End Property

Public Property Get SelectedItems() As Collection
End Property

Public Property Get Self() As IExplorerController
End Property

Public Sub Display()
End Sub
