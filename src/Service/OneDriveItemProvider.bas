VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneDriveItemProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Service")
Option Explicit

Implements IItemProvider

Private fileFactory As IFileFactory
Private folderFactory As IFolderFactory

Public Property Get Self() As OneDriveItemProvider
    Set Self = Me
End Property
Private Property Get IItemProvider_Self() As IItemProvider
    Set IItemProvider_Self = Self
End Property

Public Sub Init(ByRef cFileFactory As IFileFactory, ByRef cFolderFactory As IFolderFactory)
    GuardClauses.IsNothing cFileFactory, TypeName(cFileFactory)
    GuardClauses.IsNothing cFolderFactory, TypeName(cFolderFactory)
    Set fileFactory = cFileFactory
    Set folderFactory = cFolderFactory
End Sub

Public Function GetItem(ByVal path As String, Optional ByRef Parent As IDriveItem) As IDriveItem

    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItem"
    
    ' TODO
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Function
Private Function IItemProvider_GetItem(ByVal path As String, Optional ByRef Parent As IDriveItem) As IDriveItem
    Set IItemProvider_GetItem = GetItem(path, Parent)
End Function

Public Function GetItems(ByRef Parent As IDriveItem) As Collection

    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItems"
    
    ' TODO
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Function
Private Function IItemProvider_GetItems(ByRef Parent As IDriveItem) As Collection
    Set IItemProvider_GetItems = GetItems(Parent)
End Function

Private Function FsoFileToOneDriveFile(ByRef fso As Scripting.FileSystemObject, ByVal path As String, ByRef Parent As IDriveItem) As IFile
    Dim item As file
    Set item = fso.GetFile(path)
    Set FsoFileToOneDriveFile = fileFactory.NewFile(item.path, item.name, item.DateLastModified, item.DateCreated, item.Size, Parent, item.path)
End Function

Private Function FsoFolderToOnedriveFolder(ByRef fso As Scripting.FileSystemObject, ByVal path As String, ByRef Parent As IDriveItem) As IFolder
    Dim item As Folder
    Set item = fso.GetFolder(path)
    Set FsoFolderToOnedriveFolder = folderFactory.NewFolder(item.path, item.name, Parent, 0, item.path, Me)
End Function



