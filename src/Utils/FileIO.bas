Attribute VB_Name = "FileIO"
'@Folder("Utils")
Option Explicit

' Modu� zawiera funkcje i procedury do manipulowania plikami/ folderami
'
' Wymagania:
' - Microsoft VBScript Regular Expressions 5.5

Private fso As Object

Public Function GetTempFolderPath() As String
    
    ' Zwraca sciezke do folderu tymczasowego
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    GetTempFolderPath = fso.GetSpecialFolder(2)
    
End Function

Public Function GetCurrentFolderPath() As String
    
    ' Zwraca sciezke do aktualnego folderu
    
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    
    Dim obj As Object
    Set obj = Application
    
    Select Case obj.Name
    Case "Microsoft Excel"
        GetCurrentFolderPath = obj.ThisWorkbook.path
    Case "Microsoft Access"
        GetCurrentFolderPath = obj.CurrentProject.path
    Case "Microsoft Word"
        GetCurrentFolderPath = obj.thisdocument.path
    Case "Microsoft Outlook"
        'TODO
    End Select
    
End Function

Public Function GetNextFilename(ByVal path As String, Optional ByVal lngCounter As Long = 1, Optional ByVal strMask As String = " (i)") As String
    
    ' Zwraca piewsz� dost�pn� nazw� pliku (tzn. tak�, kt�ra nie jest zaj�ta)
    ' path -sciezka do pliku
    ' lngCounter - liczba, od kt�rej ma zacz�� liczy� (opcjonalnie)
    ' strMask - maska numeracji (opcjonalnie)
    
    Dim strFolder As String
    Dim strFileName As String
    Dim strExtension As String
    Dim strPattern As String
    Dim i As Integer
    Dim re As RegExp
    
    Set re = New RegExp
    
    For i = 1 To Len(strMask)
        strPattern = strPattern & "\" & Mid(strMask, i, 1)
    Next i
    
    strPattern = Replace(strPattern, "\i", "(\d+)")
    re.pattern = strPattern
    
    Do While FileExists(path)
        If strFolder = "" Then strFolder = GetFolder(path)
        If strExtension = "" Then strExtension = GetFileExtension(path)
        If strFileName = "" Then
            strFileName = GetFileName(path)
            strFileName = Left(strFileName, Len(strFileName) - (Len(strExtension) + 1))
        End If
        
        ' Sprawdzanie, czy nazwa pliku ju� zosta�a poddana temu "zabiegowi"
        If re.Test(strFileName) Then
            strFileName = re.Replace(strFileName, "")
        End If
        
        path = FileIO.CombinePath(strFolder, strFileName & Replace(strMask, "i", CStr(lngCounter)) & "." & strExtension)
        lngCounter = lngCounter + 1
    Loop
    
    GetNextFilename = path
    
End Function

Public Function FileExists(ByVal path As String) As Boolean

    ' Sprawdza, czy plik istnieje
    ' path - sciezka do pliku

    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    
    If fso.FileExists(path) = True Then
        FileExists = True
    End If

End Function

Public Function FolderExists(ByVal path As String) As Boolean

    ' Sprawdza, czy folder istnieje
    ' path - sciezka do folderu

    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    
    If fso.FolderExists(path) = True Then
        FolderExists = True
    End If

End Function

Public Function CreateFolder(ByVal path As String) As Boolean
   
    ' Tworzy podany katalog
    ' path - sciezka do katalogu
    
    On Error Resume Next
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    
    fso.CreateFolder (path)
    
    If FolderExists(path) Then
        CreateFolder = True
    End If
    
    Exit Function

End Function

Public Function BuildPath(ByVal path As String) As Boolean

    ' Tworzy wszystkie foldery w podanej �cie�e, je�eli nie istniej�

    On Error GoTo ErrHandler
    
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    
    Dim paths As Variant
    paths = SplitPath(path)
    
    Dim pathsMerged As Variant
    Dim tempPath As String
    
    Dim i As Integer
    pathsMerged = Array(paths(LBound(paths)))
    For i = LBound(paths) + 1 To UBound(paths)
        pathsMerged = Utils.ArrayAddItem(pathsMerged, paths(i))
        
        tempPath = FileIO.CombinePath(pathsMerged)
        If Not FileIO.FolderExists(tempPath) And Not FileIO.FileExists(tempPath) Then
            If Not FileIO.CreateFolder(tempPath) Then
                Exit Function
            End If
        End If
    Next i
    
    BuildPath = True
    
    Exit Function
    
ErrHandler:
    err.Clear

End Function

Public Sub DeleteFile(ByVal path As String)
    
    ' Usuwa plik
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If

    If FileIO.FileExists(path) Then
        Call fso.DeleteFile(path, True)
    End If

End Sub

Public Function GetFileExtension(ByVal path As String) As String

    ' Zwraca rozszerzenie pliku.
    ' path - sciezka do pliku
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If

    GetFileExtension = fso.GetExtensionName(path)

End Function

Public Function GetFolder(ByVal path As String) As String

    ' Zwraca sciezka do folderu, w ktorym jest plik
    ' path - sciezka do pliku

    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If

    GetFolder = fso.GetParentFolderName(path)

End Function

Public Function GetFileName(ByVal path As String) As String

    ' Zwraca nazwe pliku
    ' path - sciezka do pliku

    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If

    GetFileName = fso.GetFileName(path)

End Function

Public Function GetFilesRecursive(ByVal path As String, Optional ByVal intMaxLevel As Integer = -1, Optional ByVal intLevel As Integer = 0, Optional ByRef fso As Object) As Variant

    ' Zwraca list� plik�w w folderze i subfolderach
    ' path - sciezka do pliku
    ' intMaxLevel - maksymalna ilo�� podfolder�w, z kt�rych ma pobra� pliki (-1 = bez limitu = domyslna warto��)
    ' intLevel - aktualny poziom podfolderu; nie powinno sie wywo�ywa� funkcji z inn�, ni�
    ' domy�lna, warto�ci� argumentu
    ' FSO -
    ' objFolder -

    Dim objFolder As Object
    Dim varFile As Variant
    Dim varFolder As Variant
    Dim varElements As Variant
    Dim var As Variant
    Dim i As Long
  
    On Error GoTo ErrHandler
    If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")
    Set objFolder = fso.GetFolder(path)
    
    If objFolder.files.count > 0 Then
        For Each varFile In objFolder.files
            If varFile.path <> ThisWorkbook.Fullname Then
                If IsEmpty(var) Then
                    ReDim var(0)
                Else
                    ReDim Preserve var(UBound(var) + 1)
                End If
            
                var(UBound(var)) = varFile.path
                i = i + 1
            End If
        Next varFile
        
        intLevel = intLevel + 1
    End If
    
    If objFolder.Subfolders.count > 0 Then
        If intLevel <= intMaxLevel Or intMaxLevel = -1 Then
            For Each varFolder In objFolder.Subfolders
                varElements = GetFilesRecursive(varFolder.path, intMaxLevel, intLevel, fso)
                
                If Not (IsEmpty(varElements)) Then
                    For i = 0 To UBound(varElements)
                        If IsEmpty(var) Then
                            ReDim var(0)
                        Else
                            ReDim Preserve var(UBound(var) + 1)
                        End If
                        
                        var(UBound(var)) = varElements(i)
                    Next i
                End If
                
                varElements = Empty
            Next varFolder
        End If
    End If

    GetFilesRecursive = var
    
    Exit Function
    
ErrHandler:
    Select Case err.Number
    Case Is = 0
    Case Is = 70
        ' brak dost�pu do folderu
        err.Clear
    Case Else
        err.Raise (err.Number)
    End Select

End Function

Public Function GetSubfolders(ByVal path As String) As Variant

    ' Zwraca list� podfoler�w w folderze
    ' path - sciezka do pliku

    Dim objFolder As Object
    Dim varFolder As Variant
    Dim varElements As Variant
    Dim var As Variant
    Dim i As Integer
  
    On Error GoTo ErrHandler
    
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    Set objFolder = fso.GetFolder(path)
    
    If objFolder.Subfolders.count > 0 Then
        For Each varFolder In objFolder.Subfolders
            If IsEmpty(var) Then
                ReDim var(0)
            Else
                ReDim Preserve var(UBound(var) + 1)
            End If
            
            var(UBound(var)) = varFolder.path
            varElements = GetSubfolders(varFolder.path)
            
            If Not (IsEmpty(varElements)) = True Then
                For i = 0 To UBound(varElements)
                    If IsEmpty(var) Then
                        ReDim var(0)
                    Else
                        ReDim Preserve var(UBound(var) + 1)
                    End If
                    
                    var(UBound(var)) = varElements(i)
                Next i
            End If
            
            Set varElements = Nothing
        Next varFolder
    End If

    GetSubfolders = var
    
    Exit Function
    
ErrHandler:
    Select Case err.Number
    Case Is = 0
    Case Is = 70
        ' brak dost�pu do folderu
        err.Clear
    Case Else
        err.Raise (err.Number)
    End Select

End Function

Public Function IsFileOpen(path As String) As Boolean

    ' Sprawdza, czy podany plik jest otwarty przez inny program/ proces
    ' path - sciezka do pliku

    Dim iFilenum As Long
    Dim iErr As Long
     
    On Error Resume Next
    iFilenum = FreeFile()
    Open path For Input Lock Read As #iFilenum
    Close iFilenum
    iErr = err
    On Error GoTo 0
    
    Select Case iErr
    Case 0:    IsFileOpen = False
    Case 70:   IsFileOpen = True
    Case Else: Error iErr
    End Select
     
End Function

Public Function IsFolderReadOnly(ByVal strFolderPath As String) As Boolean

    ' Sprawdza, czy mo�na zapisa� plik w podanej �cie�ce - czy podany folder jest tylko do odczytu
    ' strFolderPath - sciezka do folderu
    
    Dim tempPath As String
    tempPath = FileIO.CombinePath(strFolderPath, "83cf4bcc-f8b3-4554-93db-b8e210cba8e0.txt")
    IsFolderReadOnly = CreateFile(tempPath, True)
     
End Function

Public Function CreateFile(ByVal path As String, Optional ByVal blnDeleteFileAfter As Boolean = True)

    ' Tworzy pusty plik w podanej �cie�ce i zwraca informacj�, czy stworzenie pliku si� powiod�o
    ' path -sciezka do pliku
    ' blnDeleteFileAfter - okre�la, czy skasowa� plik po zapisaniu go na dysku
    
    On Error GoTo ErrHandler
    
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If

    fso.CreateTextFile (path)
    
    If fso.FileExists(path) = True Then
        CreateFile = True
        If blnDeleteFileAfter Then fso.DeleteFile (path)
    End If
    
    Exit Function

ErrHandler:
    err.Clear

End Function

Public Sub WriteFile(ByVal path As String, _
                    ByVal value As String, _
                    ByVal append As Boolean, _
                    Optional ByVal openAfterWrite As Boolean = False)

    ' Zapisuje podany string do pliku. Procedura domy�lnie tworzy plik, je�eli plik nie istnieje
    ' path - sciezka do pliku
    ' strValue - warto�� do zapisania do pliku
    ' blnAppend - okre�la, czy ma dopisa� do pliku, czy go nadpisa�

    Dim textStream As Object
    Dim streamMode As Integer
    
    If append = True Then
        streamMode = 8
    Else
        streamMode = 2
    End If
    
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If

    'Open the text file
    Set textStream = fso.OpenTextFile(path, streamMode, True)
    
    'Display the contents of the text file
    Call textStream.Write(value)
    
    'Close the file and clean up
    textStream.Close
    
    If openAfterWrite Then
        Shell "Notepad " & path, vbNormalFocus
    End If

End Sub

Public Sub WriteFileAlt(ByVal path As String, _
                        ByVal value As String, _
                        Optional ByVal charset As String = "UTF-8", _
                        Optional ByVal append As Boolean = False, _
                        Optional ByVal openAfterWrite As Boolean = False)

    ' Mo�e by� przydatne do zamiany utf-8 na win-1250 i vice-versa
    ' Alternatywna procedura do zapisu stringu do pliku. W tej procedurze mo�na ustali� kodowanie znak�w oraz to, czy plik ma zosta� otworzony po
    ' zapisie
    ' path - sciezka do pliku
    ' strValue - warto�� do zapisania do pliku
    ' strCharset - kodowanie znak�w
    ' blnAppend - okre�la, czy ma dopisa� do pliku, czy go nadpisa�
    ' blnOpenAfterWrite - okre�la, czy otworzy� plik po zapisaniu na dysku
    
    ' Dost�pne charesty:
    ' Sets or returns a String value that specifies the character set into which the contents of the Stream will be translated.
    ' The default value is Unicode. Allowed values are typical strings passed over the interface as Internet character set names (for example, "iso-8859-1", "Windows-1252", and so on).
    ' For a list of the character set names that are known by a system, see the subkeys of HKEY_CLASSES_ROOT\MIME\Database\Charset in the Windows Registry.
    
    On Error GoTo ErrHandler
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 'Specify stream type - we want To save text/string data.
    stream.charset = charset 'Specify charset For the source text data.
    stream.Open
    
    If append And FileIO.FileExists(path) Then
        value = ReadFileAlt(path, charset) & vbCrLf & value
    End If
    
    Call stream.WriteText(value)
    Call stream.SaveToFile(path, 2) 'Save binary data To disk
    Call stream.Close
    
    If openAfterWrite = True Then
        Call FileIO.OpenFile("notepad", path)
    End If
    
    Exit Sub
    
ErrHandler:

    'Close the file and clean up
    Call stream.Close
    
    Select Case err.Number
    Case Is = 0
    Case Else
        err.Raise err.Number, "FileIO::ReadFileAlt()", err.Description
    End Select
    
End Sub

Public Sub WriteJson(ByVal path As String, ByRef dictOrCollection As Object)

    ' Zamienia podany obiekt s�ownik/ kolekcj� do formatu JSON i zapisuje
    ' do wskazanego pliku path
    Dim json As String
    Set dictOrCollection = IIf(dictOrCollection Is Nothing, New Collection, dictOrCollection)
    json = JsonConverter.ConvertToJson(dictOrCollection)
    Call FileIO.WriteFileAlt(path, json)

End Sub

Public Sub WriteBinaryFile(ByVal path As String, ByVal varByteArray As Variant)

    ' Zapisuje podan� macierz bajt�w do pliku
    ' path -sciezka do pliku
    ' varByteArray - macierz bajt�w

    On Error GoTo ErrHandler
    
    Dim stream As ADODB.stream
    Set stream = New ADODB.stream
    stream.Type = adTypeBinary
    stream.Open
    stream.Write varByteArray
    stream.SaveToFile path, adSaveCreateNotExist
    stream.Close
    
    Exit Sub
    
ErrHandler:
    stream.Close
    err.Raise err.Number, err.Source & "FileIO.WriteBinaryFile", err.Description

End Sub

Public Function ReadFile(ByVal path As String) As String

    ' Wczytuje w ca�o�ci podany plik
    ' path -sciezka do pliku

    Dim fso As Object
    Dim objFile As Object
    Dim objReadFile As Object
    
    On Error GoTo ErrHandler
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If

    'Open the text file
    Set objFile = fso.GetFile(path)
    
    If objFile.Size > 0 Then
        Set objReadFile = fso.OpenTextFile(path, 1)
        ReadFile = objReadFile.ReadAll
        objReadFile.Close
    End If
    
    Exit Function
    
ErrHandler:

    'Close the file and clean up
    Select Case err.Number
    Case Is = 0
    Case Else
        err.Raise err.Number, "FileIO::ReadFile()", err.Description
    End Select

End Function

Public Function ReadFileAlt(ByVal path As String, ByVal charset As String, Optional ByVal numchars As Long = -1) As String

    ' Alternatywna procedura do wczytywania pliku do stringu. W tej procedurze mo�na
    ' ustali� kodowanie znak�w otwieranego pliku
    ' path - sciezka do pliku
    ' strCharset - kodowanie znak�w
    ' Dost�pne charesty:
    ' Sets or returns a String value that specifies the character set into which the contents of the Stream will be translated.
    ' The default value is Unicode. Allowed values are typical strings passed over the interface as Internet character set names (for example, "iso-8859-1", "Windows-1252", and so on).
    ' For a list of the character set names that are known by a system, see the subkeys of HKEY_CLASSES_ROOT\MIME\Database\Charset in the Windows Registry.
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 'Specify stream type - we want To save text/string data.
    stream.charset = charset 'Specify charset For the source text data.
    stream.Open
    stream.LoadFromFile path
    ReadFileAlt = stream.ReadText(numchars)
    Call stream.Close
    
    Exit Function
    
ErrHandler:

    'Close the file and clean up
    Call stream.Close
    
    Select Case err.Number
    Case Is = 0
    Case Else
        err.Raise err.Number, "FileIO::ReadFileAlt()", err.Description
    End Select
    
End Function

Public Sub OpenFile(ByVal strAppName As String, ByVal path As String, Optional ByVal vbWindowStyle As VBA.VbAppWinStyle = vbNormalFocus)

    ' Otwiera podany plik z uzyciem podanej aplikacji
    ' strAppName - nazwa aplikacji (np. notepad, excel, winword, firefox)
    ' path - sciezka do pliku
    ' vbWindowStyle - tryb w jakim ma otworzyc plik (pelnoekranowy,
    ' zminimalizowany, itp.)
    
    On Error GoTo ErrHandler
    Shell strAppName & " " & path, vbWindowStyle
    Exit Sub
    
ErrHandler:
    err.Raise err.Number, "FileIO::OpenFile()", err.Description

End Sub

Public Sub OpenFileAlt(ByVal path As String)
    ThisWorkbook.FollowHyperlink path
End Sub

Public Function ChooseFolder(Optional ByVal strInitialFolder As String = "", Optional ByVal strTitle As String = "Wybierz folder") As String

    ' Funkcja do wyboru folderu
    ' strInitialFolder - folder startowy (opcjonalnie)
    ' strTitle - tytu� okna dialogowego (opcjonalnie)
    
    Dim fldr As FileDialog
    Dim strElement As String
    
    If strInitialFolder = "" Then strInitialFolder = ThisWorkbook.path
    
    Set fldr = Application.FileDialog(4)
    With fldr
        .title = strTitle
        .InitialFileName = strInitialFolder
        If .Show <> -1 Then GoTo NextCode
        strElement = .SelectedItems(1)
    End With
    
NextCode:
    ChooseFolder = strElement
    Set fldr = Nothing
    
End Function

Public Function ChooseFiles(Optional ByVal strInitialFolder As String = "", Optional ByVal strTitle As String = "Wybierz plik", Optional ByVal blnAllowMultiSelect As Boolean = False, Optional ByVal colFilters As Collection) As Variant

    ' Funkcja do wyboru pliku/ plikďż˝w
    ' strInitialFolder - folder startowy (opcjonalnie)
    ' strTitle - tytuďż˝ okna dialogowego (opcjonalnie)
    ' blnAllowMultiSelect - okreďż˝la, czy moďż˝na wybrac wiďż˝cej, niďż˝ jeden plik (opcjonalnie)
    '
    ' Przykład, jak zrobić filtry:
    ' Dim d As Object
    ' Set d = CreateObject("Scripting.Dictionary")
    ' d.Add "extensions", "*.csv"
    ' d.Add "description", "Pliki CSV"
    ' Dim filters As Collection
    ' Set filters = New Collection
    ' filters.Add d
    
    Dim fldr As FileDialog
    Dim var As Variant
    Dim i As Integer
    
    If strInitialFolder = "" Then strInitialFolder = ThisWorkbook.path
    
    Set fldr = Application.FileDialog(3)
    
    With fldr
        .title = strTitle
        .AllowMultiSelect = blnAllowMultiSelect
        .InitialFileName = strInitialFolder
        .filters.Clear
        
        If Not colFilters Is Nothing Then
            If colFilters.count > 0 Then
                Dim d As Object
                For Each d In colFilters
                    If d.Exists("position") Then
                        Call .filters.Add(d("description"), d("extensions"), d("position"))
                    Else
                        Call .filters.Add(d("description"), d("extensions"))
                    End If
                    
                Next d
            End If
        Else
            Call .filters.Add("Wszystkie pliki", "*.*", 1)
        End If
        
        If .Show <> -1 Then GoTo NextCode
        If .SelectedItems.count = 1 Then
            var = Array(.SelectedItems(1))
        Else
            For i = 1 To .SelectedItems.count
                If IsEmpty(var) Then
                    ReDim var(0)
                Else
                    ReDim Preserve var(UBound(var) + 1)
                End If
                
                var(UBound(var)) = .SelectedItems(i)
            
                ' Call ArrayAddItemSub(varElements, .SelectedItems(i))
            Next i
        End If
    End With
    
NextCode:
    ChooseFiles = var
    Set fldr = Nothing
    
End Function

Public Function ChooseFileToSave(Optional ByVal InitialFolder As String = "", _
                                Optional ByVal title As String = "Zapisz jako", _
                                Optional ByVal filter As String) As String

    ' Funkcja do wyboru pliku/ plik�w
    ' strInitialFolder - folder startowy (opcjonalnie)
    ' strTitle - tytu� okna dialogowego (opcjonalnie)
    ' filter = "Text Files (*.txt), *.txt,CSV Files (*.csv), *.csv"
    
    ChooseFileToSave = Application.GetSaveAsFilename(fileFilter:=filter, _
                        InitialFileName:=InitialFolder, _
                        title:=title)
    
End Function

Public Function RemoveIllegalChars(ByVal value As String) As String

    ' Usuwa nielegalne znaki z podanego stringa. Zwraca string
    ' Nielegalne, czyli takie znaki, kt�re s� niedopuszczone w nazwie pliku
    ' /folderu
    ' strValue - string, z kt�rego maj� zosta� usuni�te nielegalne znaki
    ' zwraca string bez nielegalnych znak�w
    
    Dim arr As Variant
    arr = Array("<", ">", ":", Chr(34), "/", "\", "|", "?", "*")
    
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        value = Replace(value, arr(i), "", 1)
    Next i
    
    RemoveIllegalChars = value

End Function

Public Sub DeleteFolder(ByVal path As String)
    
    ' Usuwa podany folder, wraz z całą zawartością
    If FileIO.FolderExists(path) Then
        If fso Is Nothing Then
            Set fso = CreateObject("Scripting.FileSystemObject")
        End If
        Call fso.DeleteFolder(path)
    End If

End Sub

Public Sub ClearFolder(ByVal path As String)

    ' Usuwa wszystkie pliki i subfoldery z wybranego folderu
    Dim files As Variant
    files = FileIO.GetFilesRecursive(path)
        
    If Utils.IsArrayDimensioned(files) Then
        Dim File As Variant
        For Each File In files
            Call FileIO.DeleteFile(CStr(File))
        Next File
    End If
    
    Dim folders As Variant
    folders = FileIO.GetSubfolders(path)
    
    If Utils.IsArrayDimensioned(folders) Then
        Dim folder As Variant
        For Each folder In folders
            Call FileIO.DeleteFolder(CStr(folder))
        Next folder
    End If

End Sub

Public Function GetRandomFilePath(Optional ByVal ext As String) As String

    ' Zwraca sciezke do pliku w folderze tymczasowym o losowej nazwie
    Dim path As String
    path = Utils.GetRandomString(10)
    path = path & IIf(ext <> vbNullString, ".", "") & ext
    path = CombinePath(Environ("temp"), path)
    GetRandomFilePath = path

End Function

Public Function CombinePath(ParamArray args()) As String

    ' Łączy ...
    
    Dim args2 As Variant
    If Utils.IsArrayDimensioned(args(0)) Then
        args2 = args(0)
    Else
        args2 = args
    End If

    Dim path As String
    path = args2(0)
    
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    
    Dim i As Integer
    For i = 1 To UBound(args2)
        path = fso.BuildPath(path, args2(i))
    Next i
    
    CombinePath = path
    
End Function

Public Function SplitPath(ByVal path As String) As Variant

    ' Rozdziela podana �cie�k� wg folder�w i ew plik�w i zwraca list�
    If path <> vbNullString Then
        Dim tempPath As String
        tempPath = path
        
        Dim x As String
        Dim off As Integer
        Dim var As Variant
        
        Do While Len(tempPath) > 0
            x = FileIO.GetFileName(tempPath)
            
            If x = vbNullString Then
                Exit Do
            Else
                var = Utils.ArrayAddItem(var, x)
                off = CInt(Right(tempPath, 1) = "\") * -1
            End If
            
            tempPath = Left(tempPath, Len(tempPath) - (Len(x) + off))
            'Debug.Print tempPath
        Loop
        
        If Right(tempPath, 1) = "\" Then
'            tempPath = Left(tempPath, Len(tempPath) - 1)
        End If
        
        var = Utils.ArrayAddItem(var, tempPath)
        
        Dim var2 As Variant
        Dim i As Integer
        For i = UBound(var) To LBound(var) Step -1
            var2 = Utils.ArrayAddItem(var2, var(i))
        Next i
        
        SplitPath = var2
    End If

End Function

Public Function FileSize(ByVal path As String) As Long

    ' Zwraca wielkośc pliku w bajtach
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If

    Dim f As Object
    Set f = fso.GetFile(path)
    FileSize = f.Size

End Function

Public Sub MoveFile(ByVal Source As String, ByVal dest As String)

    ' Przenosi/ zmienia nazwę pliku z source na dest
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.MoveFile(Source, dest)

End Sub

Public Sub CopyFile(ByVal Source As String, ByVal dest As String, Optional ByVal overwrite As Boolean = False)

    ' Przenosi/ zmienia nazwę pliku z source na dest
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.CopyFile(Source, dest, overwrite)

End Sub

Public Function ReplaceFileExt(ByVal path As String, ByVal ext As String) As String

    ' Zamienia rozszerzenie pliku path na ext; np. C:\Test\output.txt ->
    ' C:\test\output.log
    ' path - nazwa/ ściezka do pliku, którego rozszerzenie ma zostać zamienione
    ' ext - nowe rozszerzenie

    Dim oldExt As String
    oldExt = FileIO.GetFileExtension(path)
    
    If oldExt <> vbNullString Then
        Dim extLen As Integer
        extLen = Len(oldExt) + 1
        path = Mid$(path, 1, Len(path) - extLen)
    End If
    
    If Right(path, 1) <> "." Then
        path = path & IIf(ext <> vbNullString, ".", "") & ext
    Else
        path = path & ext
    End If
    
    ReplaceFileExt = path

End Function
