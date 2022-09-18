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

Public Function GetItemById(ByVal Id As String, Optional ByRef parent As IDriveItem) As IDriveItem

    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItemById"
    
    Dim DriveId As String
    If Not parent Is Nothing Then DriveId = parent.DriveId
    
    Dim json As String
    json = api.GetItemById(Id, DriveId)
    
    Dim item As IDriveItem
    Set item = JsonToIDriveItem(json, parent)
    Set GetItemById = item
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source & ";" & Self, Err.Description

End Function
Private Function IItemProvider_GetItemById(ByVal Id As String, Optional ByRef parent As IDriveItem) As IDriveItem
    Set IItemProvider_GetItemById = GetItemById(Id, parent)
End Function

Public Function GetItemByPath(ByVal path As String, Optional ByRef parent As IDriveItem) As IDriveItem

    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItemByPath"
    
    Dim json As String
    json = api.GetItemByPath(path)
    
    Dim item As IDriveItem
    Set item = JsonToIDriveItem(json, parent)
    Set GetItemByPath = item
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source & ";" & Self, Err.Description

End Function
Private Function IItemProvider_GetItemByPath(ByVal path As String, Optional ByRef parent As IDriveItem) As IDriveItem
    Set IItemProvider_GetItemByPath = GetItemByPath(path, parent)
End Function

Public Function GetItems(ByRef parent As IDriveItem) As Collection

    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetItems"
    
    Dim json As String
    json = api.GetItems(parent)
    
    Dim items As Collection
    Set items = JsonToIDriveItems(json, parent)
    Set GetItems = items
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source & ";" & Self, Err.Description

End Function
Private Function IItemProvider_GetItems(ByRef parent As IDriveItem) As Collection
    Set IItemProvider_GetItems = GetItems(parent)
End Function

Private Function JsonToIDriveItem(ByVal json As String, ByRef parent As IDriveItem) As IDriveItem
    
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".JsonToIDriveItem"
    
    Dim dict As Scripting.Dictionary
    If TryParseJson(json, dict) Then
        Dim item As IDriveItem
        Set item = IDriveItemFromDictionary(dict, parent)
        Set JsonToIDriveItem = item
        
    Else
        Err.Raise ErrorCodes.JsonParseError, Self, "Bad json response"
        ' TODO: log bad json
    
    End If
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source & ";" & Self, Err.Description
    
End Function

Private Function JsonToIDriveItems(ByVal json As String, ByRef parent As IDriveItem) As Collection
    
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".JsonToIDriveItems"
    
    Dim dict As Scripting.Dictionary
    If TryParseJson(json, dict) Then
        Dim resultCol As Collection
        Set resultCol = New Collection
        
        Dim item As IDriveItem
        
        Dim col As Collection
        Set col = dict("value")
        If Not col Is Nothing Then
            Dim d As Scripting.Dictionary
            For Each d In col
                Set item = IDriveItemFromDictionary(d, parent)
                If Not item Is Nothing Then
                    resultCol.Add item
                End If
            Next d
        End If
        
        Set JsonToIDriveItems = resultCol
    
    Else
        Err.Raise ErrorCodes.JsonParseError, Self, "Bad json response"
        ' TODO: log bad json
    
    End If
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source & ";" & Self, Err.Description
    
End Function

Private Function IDriveItemFromDictionary(ByRef dict As Scripting.Dictionary, ByRef parent As IDriveItem) As IDriveItem
    
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".IDriveItemFromDictionary"
    
    If Not dict Is Nothing Then
        If dict.Exists("folder") Then
            Dim fld As IFolder
            Set fld = folderFactory.NewFolderFromDictionary(dict, parent, Me)
            Set IDriveItemFromDictionary = fld
            
        ElseIf dict.Exists("file") Then
            Dim fle As IFile
            Set fle = fileFactory.NewFileFromDictionary(dict, parent)
            Set IDriveItemFromDictionary = fle
            
        Else
            ' This could be OneNote notebook
            'err.Raise ErrorCodes.BadIDriveItemDictionary, Self, "Dictionary item 'file' or 'folder' not found"
        
        End If
    Else
        Err.Raise ErrorCodes.BadIDriveItemDictionary, Self, "Dictionary cannot be nothing"
        
    End If
    
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, Err.Source & ";" & Self, Err.Description
    
End Function

Private Function TryParseJson(ByVal json As String, ByRef obj As Object) As Boolean
    On Error GoTo ErrHandler
    Set obj = ParseJson(json)
    TryParseJson = True
    Exit Function
ErrHandler:
    Err.Clear
End Function
