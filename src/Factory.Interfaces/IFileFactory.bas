VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFileFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Factory.Interfaces")
'@Interface
Option Explicit

Public Property Get Self() As IFileFactory
End Property

Public Function NewFile(ByVal Id As String, _
                        ByVal DriveId As String, _
                        ByVal Name As String, _
                        ByVal LastModifiedTime As Date, _
                        ByVal CreatedTime As Date, _
                        ByVal Size As String, _
                        ByRef Parent As IDriveItem, _
                        ByVal Path As String) As IFile
End Function

Public Function NewFileFromDictionary(ByRef dict As Scripting.Dictionary, ByRef Parent As IDriveItem) As IFile
End Function




