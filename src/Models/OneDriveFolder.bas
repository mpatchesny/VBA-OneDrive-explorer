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
    json As String
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

Public Property Get Name() As String
    Name = this.Name
End Property
Private Property Get IFolder_Name() As String
    IFolder_Name = Name
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

Public Property Get Parent() As IDriveItem
    ' TODO
    Set Parent = Nothing
End Property
Private Property Get IDriveItem_Parent() As IDriveItem
    Set IDriveItem_Parent = Parent
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


Public Property Get ChildrenCount() As Long
    Dim Self As String
    Self = TypeName(Me) & ".ChildrenCount"
    ' TODO
End Property
Private Property Get IFolder_ChildrenCount() As Long
    IFolder_ChildrenCount = ChildrenCount
End Property

Public Function GetChildren() As Collection
    Dim Self As String
    Self = TypeName(Me) & ".GetChildren"
    ' TODO
    Set GetChildren = Nothing
End Function
Private Function IFolder_GetChildren() As Collection
    Set IFolder_GetChildren = GetChildren
End Function



