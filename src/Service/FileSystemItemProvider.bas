VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemItemProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Service")
Option Explicit

Implements IItemProvider

Public Function GetItem(ByVal path As String, Optional ByRef parent As IDriveItem) As IDriveItem

    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItem"
    
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim item As IDriveItem
    
    If fso.FolderExists(path) Then
        Dim fldr As IFolder
        Set fldr = FsoFolderToOnedriveFolder(fso, path, parent)
        Set item = fldr
    
    ElseIf fso.FileExists(path) Then
        Dim fle As IFile
        Set fle = FsoFileToOneDriveFile(fso, path)
        Set item = fle
        
    Else
        err.Raise ErrorCodes.PathAccessError, Self, "Path not found or is inaccessible : " & path
        
    End If
    
    Set GetItem = item
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Function
Private Function IItemProvider_GetItem(ByVal path As String, Optional ByRef parent As IDriveItem) As IDriveItem
    Set IItemProvider_GetItem = GetItem(path, parent)
End Function


Public Function GetItems(ByRef parent As IDriveItem) As Collection

    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItems"
    
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim thisFolder As Folder
    Set thisFolder = fso.GetFolder(parent.path)
    
    Dim col As Collection
    Set col = New Collection
    
    Dim item2 As Folder
    For Each item2 In thisFolder.SubFolders
        col.Add FsoFileToOneDriveFile(fso, item2.path)
    Next item2

    Dim item As file
    For Each item In thisFolder.files
        col.Add FsoFolderToOnedriveFolder(fso, item.path, parent)
    Next item
    
    Set GetItems = col
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Function
Private Function IItemProvider_GetItems(ByRef parent As IDriveItem) As Collection
    Set IItemProvider_GetItems = GetItems(parent)
End Function

Private Function FsoFileToOneDriveFile(ByRef fso As Scripting.FileSystemObject, ByVal path As String) As IFile
    Dim item As file
    Set item = fso.GetFile(path)
    
    With New OneDriveFile
        .Init item.path, item.Name, item.DateLastModified, item.DateCreated, item.Size, New FileSystemItemProvider, item.path
        Set FsoFileToOneDriveFile = .Self
    End With
End Function

Private Function FsoFolderToOnedriveFolder(ByRef fso As Scripting.FileSystemObject, ByVal path As String, ByRef parent As IDriveItem) As IFolder
    Dim item As Folder
    Set item = fso.GetFolder(path)
    
    With New OneDriveFolder
        .Init item.path, item.Name, parent, 0, item.path, New FileSystemItemProvider
        Set FsoFolderToOnedriveFolder = .Self
    End With
End Function

