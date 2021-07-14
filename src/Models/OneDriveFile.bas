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
    DriveId As String
    Name As String
    LastModifiedTime As Date
    CreatedTime As Date
    Size As String
    Parent As IDriveItem
    Path As String
    DownloadUrl As String
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
Private Property Let Id(ByVal value As String)
    this.Id = value
End Property

Public Property Get DriveId() As String
    DriveId = this.DriveId
End Property
Private Property Get IDriveItem_DriveId() As String
    IDriveItem_DriveId = DriveId
End Property
Private Property Get IFile_DriveId() As String
    IFile_DriveId = DriveId
End Property
Private Property Let DriveId(ByVal value As String)
    this.DriveId = value
End Property

Public Property Get Name() As String
    Name = this.Name
End Property
Private Property Get IFile_Name() As String
    IFile_Name = Name
End Property
Private Property Let Name(ByVal value As String)
    this.Name = value
End Property

Private Property Get CreatedTime() As Date
    CreatedTime = this.CreatedTime
End Property
Private Property Get IFile_CreatedTime() As Date
    IFile_CreatedTime = CreatedTime
End Property
Private Property Let CreatedTime(ByVal value As Date)
    this.CreatedTime = value
End Property

Public Property Get LastModifiedTime() As Date
    LastModifiedTime = this.LastModifiedTime
End Property
Private Property Get IFile_LastModifiedTime() As Date
    IFile_LastModifiedTime = LastModifiedTime
End Property
Private Property Let LastModifiedTime(ByVal value As Date)
    this.LastModifiedTime = value
End Property

Public Property Get Size() As String
    Size = this.Size
End Property
Private Property Get IFile_Size() As String
    IFile_Size = Size
End Property
Private Property Let Size(ByVal value As String)
    this.Size = value
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
    Set Parent = this.Parent
End Property
Private Property Get IDriveItem_Parent() As IDriveItem
    Set IDriveItem_Parent = Parent
End Property
Private Property Let Parent(ByVal value As IDriveItem)
    Set this.Parent = value
End Property

Public Property Get Path() As String
    Path = this.Path
End Property
Private Property Get IFile_Path() As String
    IFile_Path = Path
End Property
Private Property Get IDriveItem_Path() As String
    IDriveItem_Path = Path
End Property
Private Property Get IFolder_Path() As String
    IFolder_Path = Path
End Property
Private Property Let Path(ByVal value As String)
    this.Path = value
End Property

Public Property Get DownloadUrl() As String
    DownloadUrl = this.DownloadUrl
End Property
Private Property Let DownloadUrl(ByVal value As String)
    this.DownloadUrl = value
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

Public Sub Init(ByVal cId As String, _
                ByVal cDriveId As String, _
                ByVal cName As String, _
                ByVal cLastModifiedTime As Date, _
                ByVal cCreatedTime As Date, _
                ByVal cSize As String, _
                ByRef cParent As IDriveItem, _
                ByVal cPath As String, _
                ByVal cDownloadUrl As String)
    Id = cId
    DriveId = cDriveId
    Name = cName
    LastModifiedTime = cLastModifiedTime
    CreatedTime = cCreatedTime
    Parent = cParent
    ' if not isnumeric(csize) then
    Size = cSize
    Path = cPath
    DownloadUrl = cDownloadUrl
End Sub

