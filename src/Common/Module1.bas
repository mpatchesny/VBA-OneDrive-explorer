Attribute VB_Name = "Module1"
'@Folder("Common")
Option Explicit

Sub Test()

    On Error GoTo ErrHandler
    Dim Self As String
    Self = ".test"
    
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim path As String
    path = "C:\Users\strielok\Downloads\"
    
    Dim thisFolder As Folder
    Set thisFolder = fso.GetFolder(path)
    
    Dim col As Collection
    Set col = New Collection
    
    Dim item2 As Folder
    For Each item2 In thisFolder.SubFolders
        With New OneDriveFolder
            .Init "", item2.Name, Nothing, 0, item2.path, Nothing
            col.Add .Self
        End With
    Next item2

    Dim item As file
    For Each item In thisFolder.files
        With New OneDriveFile
            .Init "", item.Name, item.DateLastModified, item.DateCreated, item.Size, Nothing, item.path
            col.Add .Self
        End With
    Next item
    
    Debug.Print col.Count
    
    Exit Sub
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Sub
