VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IItemProviderFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Factory.Interfaces")
'@Interface
Option Explicit

Public Property Get Self() As IItemProviderFactory
End Property

Public Function NewItemProvider(ByRef fileFactory As IFileFactory, ByRef folderFactory As IFolderFactory) As IItemProvider
End Function




