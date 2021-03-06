VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IExplorerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Views.Interfaces")
'@Interface
Option Explicit

Public Enum ESelectMode
    ESelectModeAll = 1
    ESelectModeFilesOnly = 2
    ESelectModeFoldersOnly = 3
End Enum

Public Property Get Model() As IExplorerViewModel
End Property

Public Property Get IsCancelled() As Boolean
End Property

Public Property Get Self() As IExplorerView
End Property

Public Sub Display()
End Sub

Public Sub HideView()
End Sub
