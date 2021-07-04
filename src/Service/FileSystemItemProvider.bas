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
        With New OneDriveFolder
            .Init item2.path, item2.Name, parent, 0, item2.path, New FileSystemItemProvider
            col.Add .Self
        End With
    Next item2

    Dim item As file
    For Each item In thisFolder.files
        With New OneDriveFile
            .Init item2.path, item.Name, item.DateLastModified, item.DateCreated, item.Size, New FileSystemItemProvider, item.path
            col.Add .Self
        End With
    Next item
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Function
Private Function IItemProvider_GetItems(ByRef parent As IDriveItem) As Collection
    Set IItemProvider_GetItems = GetItems(parent)
End Function

