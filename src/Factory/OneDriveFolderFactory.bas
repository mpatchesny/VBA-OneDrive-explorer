VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneDriveFolderFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Factory")
Option Explicit

Implements IFolderFactory

Public Property Get Self() As OneDriveFolderFactory
    Set Self = Me
End Property
Private Property Get IFolderFactory_Self() As IFolderFactory
    Set IFolderFactory_Self = Self
End Property

Public Function newFolder(ByVal Id As String, _
                        ByVal DriveId As String, _
                        ByVal Name As String, _
                        ByRef Parent As IDriveItem, _
                        ByVal ChildrenCount As Long, _
                        ByVal Path As String, _
                        ByVal LastModifiedTime As Date, _
                        ByRef provider As IItemProvider) As IFolder
    With New OneDriveFolder
        .Init Id, DriveId, Name, Parent, ChildrenCount, Path, LastModifiedTime, provider
        Set newFolder = .Self
    End With
End Function
Private Function IFolderFactory_NewFolder(ByVal Id As String, _
                                            ByVal DriveId As String, _
                                            ByVal Name As String, _
                                            ByRef Parent As IDriveItem, _
                                            ByVal ChildrenCount As Long, _
                                            ByVal Path As String, _
                                            ByVal LastModifiedTime As Date, _
                                            ByRef provider As IItemProvider) As IFolder
    Set IFolderFactory_NewFolder = newFolder(Id, DriveId, Name, Parent, ChildrenCount, Path, LastModifiedTime, provider)
End Function

Public Function NewFolderFromDictionary(ByRef d As Scripting.IDictionary, _
                                        ByRef Parent As IDriveItem, _
                                        ByRef provider As IItemProvider) As IFolder
    
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".NewFolderFromDictionary"
    
    Dim ChildCount As Long
    If d.Exists("folder") Then
        ChildCount = d("folder")("childCount")
    End If
    
    Dim lastModified As Date
    If d.Exists("lastModifiedDateTime") Then
        lastModified = JsonConverter.ParseIso(d("lastModifiedDateTime"))
    End If
    
    Dim tempDict As Scripting.Dictionary
    Set tempDict = d
    If d.Exists("remoteItem") Then
        Set tempDict = d("remoteItem")
    End If
    
    Dim DriveId As String
    If Not tempDict Is Nothing Then
        If tempDict.Exists("parentReference") Then
            DriveId = tempDict("parentReference")("driveId")
        End If
    End If
    
    With New OneDriveFolder
        .Init d("id"), DriveId, d("name"), Parent, ChildCount, d("webUrl"), lastModified, provider
        Set NewFolderFromDictionary = .Self
    End With
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description
    
End Function
Private Function IFolderFactory_NewFolderFromDictionary(ByRef d As Scripting.IDictionary, _
                                                        ByRef Parent As IDriveItem, _
                                                        ByRef provider As IItemProvider) As IFolder
    Set IFolderFactory_NewFolderFromDictionary = NewFolderFromDictionary(d, Parent, provider)
End Function
