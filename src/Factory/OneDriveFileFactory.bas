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

Public Function NewFile(ByVal Id As String, _
                        ByVal DriveId As String, _
                        ByVal Name As String, _
                        ByVal LastModifiedTime As Date, _
                        ByVal CreatedTime As Date, _
                        ByVal Size As String, _
                        ByRef Parent As IDriveItem, _
                        ByVal Path As String) As IFile
                                    
    With New OneDriveFile
        .Init Id, DriveId, Name, LastModifiedTime, CreatedTime, Size, Parent, Path, ""
        Set NewFile = .Self
    End With

End Function
Private Function IFileFactory_NewFile(ByVal Id As String, _
                                    ByVal DriveId As String, _
                                    ByVal Name As String, _
                                    ByVal LastModifiedTime As Date, _
                                    ByVal CreatedTime As Date, _
                                    ByVal Size As String, _
                                    ByRef Parent As IDriveItem, _
                                    ByVal Path As String) As IFile
    Set IFileFactory_NewFile = NewFile(Id, DriveId, Name, LastModifiedTime, CreatedTime, Size, Parent, Path)
End Function

Public Function NewFileFromDictionary(ByRef d As Scripting.Dictionary, ByRef Parent As IDriveItem) As IFile
    
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".NewFileFromDictionary"
    
    Dim lastModified As Date
    If d.Exists("lastModifiedDateTime") Then
        lastModified = JsonConverter.ParseIso(d("lastModifiedDateTime"))
    End If
    
    Dim created As Date
    If d.Exists("createdDateTime") Then
        created = JsonConverter.ParseIso(d("createdDateTime"))
    End If
    
    Dim parentRef As Scripting.Dictionary
    Set parentRef = d("parentReference")
    If parentRef Is Nothing Then Set parentRef = New Scripting.Dictionary
    
    With New OneDriveFile
        .Init d("id"), parentRef("driveId"), d("name"), lastModified, created, d("size"), Parent, d("webUrl"), d("@microsoft.graph.downloadUrl")
        Set NewFileFromDictionary = .Self
    End With
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description
    
End Function
Private Function IFileFactory_NewFileFromDictionary(ByRef dict As Scripting.Dictionary, ByRef Parent As IDriveItem) As IFile
    Set IFileFactory_NewFileFromDictionary = NewFileFromDictionary(dict, Parent)
End Function

