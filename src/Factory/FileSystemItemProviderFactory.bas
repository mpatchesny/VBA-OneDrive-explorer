VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemItemProviderFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Factory")
Option Explicit

Implements IItemProviderFactory

Public Property Get Self() As FileSystemItemProviderFactory
    Set Self = Me
End Property
Private Property Get IItemProviderFactory_Self() As IItemProviderFactory
    Set IItemProviderFactory_Self = Self
End Property

Public Function NewItemProvider(ByRef fileFactory As IFileFactory, ByRef folderFactory As IFolderFactory) As IItemProvider
    With New FileSystemItemProvider
        .Init fileFactory, folderFactory
        Set NewItemProvider = .Self
    End With
End Function
Private Function IItemProviderFactory_NewItemProvider(fileFactory As IFileFactory, folderFactory As IFolderFactory) As IItemProvider
    Set IItemProviderFactory_NewItemProvider = NewItemProvider(fileFactory, folderFactory)
End Function

