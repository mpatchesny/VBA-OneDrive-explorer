VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneDriveFileFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Factory")
Option Explicit

Implements IFileFactory

Public Property Get Self() As OneDriveFileFactory
    Set Self = Me
End Property
Private Property Get IFileFactory_Self() As IFileFactory
    Set IFileFactory_Self = Self
End Property

Public Function NewFile(ByVal id As String, _
                                    ByVal name As String, _
                                    ByVal LastModifiedTime As Date, _
                                    ByVal CreatedTime As Date, _
                                    ByVal Size As String, _
                                    ByRef parent As IDriveItem, _
                                    ByVal path As String) As IFile
                                    
    With New OneDriveFile
        .Init id, name, LastModifiedTime, CreatedTime, Size, parent, path, ""
        Set NewFile = .Self
    End With

End Function
Private Function IFileFactory_NewFile(ByVal id As String, _
                                    ByVal name As String, _
                                    ByVal LastModifiedTime As Date, _
                                    ByVal CreatedTime As Date, _
                                    ByVal Size As String, _
                                    ByRef parent As IDriveItem, _
                                    ByVal path As String) As IFile
    Set IFileFactory_NewFile = NewFile(id, name, LastModifiedTime, CreatedTime, Size, parent, path)
End Function

Public Function NewFileFromDictionary(ByRef d As Scripting.Dictionary, ByRef parent As IDriveItem) As IFile
    
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".NewFileFromJsonString"
    
    With New OneDriveFile
        .Init d("id"), d("name"), d("lastModifiedDateTime"), d("createdDateTime"), d("size"), parent, d("webUrl"), d("@microsoft.graph.downloadUrl")
        Set NewFileFromDictionary = .Self
    End With
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description
    
End Function
Private Function IFileFactory_NewFileFromDictionary(ByRef dict As Scripting.Dictionary, ByRef parent As IDriveItem) As IFile
    Set IFileFactory_NewFileFromDictionary = NewFileFromDictionary(dict, parent)
End Function

