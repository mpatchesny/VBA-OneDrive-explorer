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
Private api As IApi

Public Property Get Self() As OneDriveItemProvider
    Set Self = Me
End Property
Private Property Get IItemProvider_Self() As IItemProvider
    Set IItemProvider_Self = Self
End Property

Public Sub Init(ByRef cFileFactory As IFileFactory, ByRef cFolderFactory As IFolderFactory, ByRef cApi As IApi)
    GuardClauses.IsNothing cFileFactory, TypeName(cFileFactory)
    GuardClauses.IsNothing cFolderFactory, TypeName(cFolderFactory)
    GuardClauses.IsNothing cApi, TypeName(cApi)
    Set fileFactory = cFileFactory
    Set folderFactory = cFolderFactory
    Set api = cApi
End Sub

Public Function GetItem(ByVal path As String, Optional ByRef parent As IDriveItem) As IDriveItem

    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItem"
    
    Dim json As String
    json = api.GetItem(path)
    
    Dim item As IDriveItem
    Set item = JsonToIDriveItem(json)
    
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
    
    ' TODO
    
    Exit Function
    
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description

End Function
Private Function IItemProvider_GetItems(ByRef parent As IDriveItem) As Collection
    Set IItemProvider_GetItems = GetItems(parent)
End Function

Private Function JsonToIDriveItem(ByVal json As String) As IDriveItem
    ' ????
    Dim item As IDriveItem
    Set JsonToIDriveItem = item
End Function


