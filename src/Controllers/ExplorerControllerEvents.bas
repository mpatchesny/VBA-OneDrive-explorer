VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExplorerControllerEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Controllers")
Option Explicit

Public Event CurrentItemChanged()

Public Property Get Self() As ExplorerControllerEvents
    Set Self = Me
End Property

Public Sub RaiseCurrentItemChanged()
    RaiseEvent CurrentItemChanged
End Sub



