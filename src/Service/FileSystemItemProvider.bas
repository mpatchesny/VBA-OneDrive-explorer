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

Private fileFactory As IFileFactory
Private folderFactory As IFolderFactory

Public Property Get Self() As FileSystemItemProvider
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

Public Function GetItemById(ByVal Id As String, Optional ByRef Parent As IDriveItem) As IDriveItem

    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItemById"
    err.Raise ErrorCodes.NotImplemented, Self, "Method GetItemById is invalid in this context"
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Function
Private Function IItemProvider_GetItemById(ByVal Id As String, Optional ByRef Parent As IDriveItem) As IDriveItem
    Set IItemProvider_GetItemById = GetItemById(Id, Parent)
End Function

Public Function GetItemByPath(ByVal Path As String, Optional ByRef Parent As IDriveItem) As IDriveItem

    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItemByPath"
    
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim item As IDriveItem
    If fso.FolderExists(Path) Then
        With FsoFolderToOnedriveFolder(fso, Path, Parent)
            Set item = .Self
        End With
    
    ElseIf fso.FileExists(Path) Then
        With FsoFileToOneDriveFile(fso, Path, Parent)
            Set item = .Self
        End With
        
    Else
        err.Raise ErrorCodes.PathAccessError, Self, "Path not found or is inaccessible : " & Path
        
    End If
    
    Set GetItemByPath = item
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Function
Private Function IItemProvider_GetItemByPath(ByVal Path As String, Optional Parent As IDriveItem) As IDriveItem
    IItemProvider_GetItemByPath = GetItemByPath(Path, Parent)
End Function

Public Function GetItems(ByRef Parent As IDriveItem) As Collection

    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItems"
    
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim thisFolder As Folder
    Set thisFolder = fso.GetFolder(Parent.Path)
    
    Dim col As Collection
    Set col = New Collection
    
    Dim item2 As Folder
    For Each item2 In thisFolder.SubFolders
        col.Add FsoFolderToOnedriveFolder(fso, item2.Path, Parent)
    Next item2

    Dim item As file
    For Each item In thisFolder.files
        col.Add FsoFileToOneDriveFile(fso, item.Path, Parent)
    Next item
    
    Set GetItems = col
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Function
Private Function IItemProvider_GetItems(ByRef Parent As IDriveItem) As Collection
    Set IItemProvider_GetItems = GetItems(Parent)
End Function

Private Function FsoFileToOneDriveFile(ByRef fso As Scripting.FileSystemObject, ByVal Path As String, ByRef Parent As IDriveItem) As IFile
    Dim item As file
    Set item = fso.GetFile(Path)
    Set FsoFileToOneDriveFile = fileFactory.NewFile(item.Path, fso.GetDrive(Path), item.Name, item.DateLastModified, item.DateCreated, item.Size, Parent, item.Path)
End Function

Private Function FsoFolderToOnedriveFolder(ByRef fso As Scripting.FileSystemObject, ByVal Path As String, ByRef Parent As IDriveItem) As IFolder
    Dim item As Folder
    Set item = fso.GetFolder(Path)
    Set FsoFolderToOnedriveFolder = folderFactory.newFolder(item.Path, fso.GetDrive(Path), item.Name, Parent, 0, item.Path, item.DateLastModified, Me)
End Function


