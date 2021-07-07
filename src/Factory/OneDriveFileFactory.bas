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
                                    ByVal Name As String, _
                                    ByVal LastModifiedTime As Date, _
                                    ByVal CreatedTime As Date, _
                                    ByVal Size As String, _
                                    ByRef Parent As IDriveItem, _
                                    ByVal path As String) As IFile
                                    
    With New OneDriveFile
        .Init Id, Name, LastModifiedTime, CreatedTime, Size, Parent, path
        Set NewFile = .Self
    End With

End Function
Private Function IFileFactory_NewFile(ByVal Id As String, _
                                    ByVal Name As String, _
                                    ByVal LastModifiedTime As Date, _
                                    ByVal CreatedTime As Date, _
                                    ByVal Size As String, _
                                    ByRef Parent As IDriveItem, _
                                    ByVal path As String) As IFile
    Set IFileFactory_NewFile = NewFile(Id, Name, LastModifiedTime, CreatedTime, Size, Parent, path)
End Function

Public Function NewFileFromJsonString(ByVal json As String) As IFile
    'TODO
End Function
Private Function IFileFactory_NewFileFromJsonString(ByVal json As String) As IFile
    Set IFileFactory_NewFileFromJsonString = NewFileFromJsonString(json)
End Function
