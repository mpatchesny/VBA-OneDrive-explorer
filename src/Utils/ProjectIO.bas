Attribute VB_Name = "ProjectIO"
'@Folder("Utils")

Option Explicit

' Procedury do exportu/ importu projektu z VBE do plików bas/frm
' Wymagane ustawienie referencji do Ms VBA extensibility library oraz dostêpu programistycznego do VBE

Private Function GetProjectPath() As String

    Dim temp As String
    temp = ThisWorkbook.path
    
    Dim folders As Variant
    folders = FileIO.SplitPath(temp)
    
    Dim folders2 As Variant
    Dim i As Integer
    If folders(UBound(folders)) = "bin" Then
        For i = 0 To UBound(folders) - 1
            folders2 = Utils.ArrayAddItem(folders2, folders(i))
        Next i
    Else
        folders2 = folders
    End If
    
    For i = 0 To UBound(folders2)
        GetProjectPath = FileIO.CombinePath(GetProjectPath, folders2(i))
    Next i
    
    GetProjectPath = FileIO.CombinePath(GetProjectPath, "src")
    
End Function

Public Sub ExportProject()

    ThisWorkbook.Save

    Dim path As String
    path = GetProjectPath
    
    FileIO.CreateFolder (path)
    ClearProjectFolder (path)

    Dim re As New RegExp
    re.pattern = "\'\@Folder\(\" & Chr(34) & "(.+)\" & Chr(34) & "\)"
    re.Global = True
    
    Dim OutputPath As String
    Dim exportFolder As String
    Dim ext As String
    Dim folder As String
    Dim reMatch As Variant
    Dim line As String
    Dim noExport As Boolean

    Dim VBProj
    Dim VBComp
    Set VBProj = Application.VBE.ActiveVBProject

    For Each VBComp In VBProj.VBComponents
        Dim i As Integer
        For i = 1 To 4
            line = VBComp.CodeModule.lines(1, i)
            
            If line Like "'@Folder(?Ignore?)" Then
                noExport = True
                Exit For
                
            ElseIf line Like "'@Folder(*)" Then
                Set reMatch = re.Execute(line)
                folder = reMatch(0).SubMatches(0)
                exportFolder = FileIO.CombinePath(path, folder)
                Call FileIO.CreateFolder(exportFolder)
                Exit For
                
            End If
        Next i
            
        If Not noExport Then
            If Len(exportFolder) = 0 Then exportFolder = path
    
            If VBComp.Type = 1 Then
                ext = ".bas"
            ElseIf VBComp.Type = 2 Then
                ext = ".bas"
            ElseIf VBComp.Type = 3 Then
                ext = ".frm"
            ElseIf VBComp.Type = 100 Then
                ext = ".doccls"
            Else
                ext = ""
            End If
    
            If ext <> "" Then
                OutputPath = FileIO.CombinePath(exportFolder, VBComp.Name & ext)
                Debug.Print ("Exporting " & VBComp.Name & " to " & OutputPath)
                VBComp.Export (OutputPath)
            End If
        End If
        
        noExport = False
    Next VBComp

End Sub

Private Sub ClearProjectFolder(ByVal path As String)

    ' Usuwanie plików *.bas, *.frm, *.frx
    If FileIO.FolderExists(path) Then
        Dim File As Variant
        Dim files As Variant
        files = FileIO.GetFilesRecursive(path)
        If IsEmpty(files) = False Then
            For Each File In files
                If FileIO.GetFileExtension(File) = "bas" Or _
                    FileIO.GetFileExtension(File) = "frm" Or _
                    FileIO.GetFileExtension(File) = "frx" Then
                    Debug.Print ("Deleting file " & File)
                    FileIO.DeleteFile (File)
                End If
            Next File
        End If
    
        ' Usuwanie pustych folderów
        Dim folder As Variant
        Dim folders As Variant
    
        folders = FileIO.GetSubfolders(path)
        If IsEmpty(folders) = False Then
            For Each folder In folders
                files = FileIO.GetFilesRecursive(folder)
    
                If IsArrayDimensioned(files) = False Then
                    Debug.Print ("Deleting empty folder " & folder)
                    FileIO.DeleteFolder (folder)
                End If
            Next folder
        End If
    End If

End Sub

Public Sub ImportProject()

    Dim path As String
    path = GetProjectPath

    Dim VBProj
    Dim VBComp
    Set VBProj = Application.VBE.ActiveVBProject
    
    Dim files As Variant
    files = FileIO.GetFilesRecursive(path)
    files = Utils.ArrayRemoveDuplicates(files)
    
    Dim moduleName As String

    Dim File As Variant
    For Each File In files
        If FileIO.GetFileExtension(File) = "bas" Or _
            FileIO.GetFileExtension(File) = "frm" Or _
            FileIO.GetFileExtension(File) = "frx" Then
            moduleName = FileIO.GetFileName(File)
            moduleName = Left(moduleName, Len(moduleName) - 4)

            For Each VBComp In VBProj.VBComponents
                If moduleName = VBComp.Name And moduleName <> "ProjectIO" Then
                    VBProj.VBComponents.Remove (VBComp)
                End If
            Next VBComp
                
            Debug.Print ("Importing " & moduleName)
            VBProj.VBComponents.Import (File)
        End If
    Next File
End Sub


