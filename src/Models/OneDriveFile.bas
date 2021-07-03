VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneDriveFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Models")
Option Explicit

Implements IDriveItem
Implements IFile

Private Type TFields
    Id As String
    Name As String
    LastModifiedTime As Date
    CreatedTime As Date
    Size As Long
    json As String
End Type
Private this As TFields

Public Property Get Id() As String
    Id = this.Id
End Property
Private Property Get IDriveItem_Id() As String
    IDriveItem_Id = Id
End Property
Private Property Get IFile_Id() As String
    IFile_Id = Id
End Property

Public Property Get Name() As String
    Name = this.Name
End Property
Private Property Get IFile_Name() As String
    IFile_Name = Name
End Property

Private Property Get CreatedTime() As Date
    CreatedTime = this.CreatedTime
End Property
Private Property Get IFile_CreatedTime() As Date
    IFile_CreatedTime = CreatedTime
End Property

Public Property Get LastModifiedTime() As Date
    LastModifiedTime = this.LastModifiedTime
End Property
Private Property Get IFile_LastModifiedTime() As Date
    IFile_LastModifiedTime = LastModifiedTime
End Property

Public Property Get Size() As Long
    Size = this.Size
End Property
Private Property Get IFile_Size() As Long
    IFile_Size = Size
End Property

Public Property Get IsFile() As Boolean
    IsFile = True
End Property
Private Property Get IDriveItem_IsFile() As Boolean
    IDriveItem_IsFile = IsFile
End Property

Public Property Get IsFolder() As Boolean
    IsFolder = False
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

Public Property Get Self() As OneDriveFile
    Set Self = Me
End Property
Private Property Get IDriveItem_Self() As IDriveItem
    Set IDriveItem_Self = Self
End Property
Private Property Get IFile_Self() As IFile
    Set IFile_Self = Self
End Property







