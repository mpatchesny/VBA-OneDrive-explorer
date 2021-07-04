VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneDriveFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Models")
Option Explicit

Implements IDriveItem
Implements IFolder

Private Type TFields
    Id As String
    Name As String
    parent As IDriveItem
    ChildrenCount As Long
    path As String
    provider As IItemProvider
End Type
Private this As TFields

Public Property Get Id() As String
    Id = this.Id
End Property
Private Property Get IDriveItem_Id() As String
    IDriveItem_Id = Id
End Property
Private Property Get IFolder_Id() As String
    IFolder_Id = Id
End Property
Private Property Let Id(ByVal value As String)
    this.Id = value
End Property

Public Property Get Name() As String
    Name = this.Name
End Property
Private Property Get IFolder_Name() As String
    IFolder_Name = Name
End Property
Private Property Let Name(ByVal value As String)
    this.Name = value
End Property

Public Property Get IsFile() As Boolean
    IsFile = False
End Property
Private Property Get IDriveItem_IsFile() As Boolean
    IDriveItem_IsFile = IsFile
End Property

Public Property Get IsFolder() As Boolean
    IsFolder = True
End Property
Private Property Get IDriveItem_IsFolder() As Boolean
    IDriveItem_IsFolder = IsFolder
End Property

Public Property Get parent() As IDriveItem
    Set parent = this.parent
End Property
Private Property Get IDriveItem_Parent() As IDriveItem
    Set IDriveItem_Parent = parent
End Property
Private Property Let parent(ByVal value As IDriveItem)
    Set this.parent = value
End Property

Public Property Get ChildrenCount() As Long
    ChildrenCount = this.ChildrenCount
End Property
Private Property Get IFolder_ChildrenCount() As Long
    IFolder_ChildrenCount = ChildrenCount
End Property
Private Property Let ChildrenCount(ByVal value As Long)
    this.ChildrenCount = value
End Property

Public Property Get path() As String
    path = this.path
End Property
Private Property Get IDriveItem_Path() As String
    IDriveItem_Path = path
End Property
Private Property Get IFile_Path() As String
    IFile_Path = path
End Property
Private Property Get IFolder_Path() As String
    IFolder_Path = path
End Property
Private Property Let path(ByVal value As String)
    this.path = value
End Property

Public Property Get Self() As OneDriveFolder
    Set Self = Me
End Property
Private Property Get IDriveItem_Self() As IDriveItem
    Set IDriveItem_Self = Self
End Property
Private Property Get IFolder_Self() As IFolder
    Set IFolder_Self = Self
End Property

Public Sub Init(ByVal cId As String, _
                ByVal cName As String, _
                ByRef cParent As IDriveItem, _
                ByVal cChildrenCount As Long, _
                ByVal cPath As String, _
                ByRef provider As IItemProvider)
    Id = cId
    Name = cName
    parent = cParent
    ChildrenCount = cChildrenCount
    path = cPath
    Set this.provider = provider
End Sub

Public Function GetChildren() As Collection
    On Error GoTo ErrHandler
    Dim Self As String
    Self = TypeName(Me) & ".GetChildren"
    
    Set GetChildren = this.provider.GetItems(Me)
    
    Exit Function
ErrHandler:
    err.Raise err.Number, err.Source & ";" & Self, err.Description
End Function
Private Function IFolder_GetChildren() As Collection
    Set IFolder_GetChildren = GetChildren
End Function
