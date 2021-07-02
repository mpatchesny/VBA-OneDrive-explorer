Attribute VB_Name = "Utils"
'@Folder("Utils")
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
#End If

' Merged module: Arrays.bas
' Merged module: Collections.bas
' Merged module: DateTime.bas
' Merged module: ExcelManipulation.bas
' Merged module: TypeConversion.bas

' =============================================================================
' Module Arrays
' =============================================================================

' Modu³ zawiera procedury i funkcje do manipulowania macierzami, zwykle typu variant
' Wymagania:
' - Microsoft VBScript Regular Expressions 5.5

Public Function IsArrayDimensioned(ByRef arr As Variant) As Boolean

    '*************************************************************
    ' Purpose:  Checks if dynamic array is dimensioned.
    ' Input:    MyArray()   - the array to examine.
    ' Output:   True - array is dimensioned; False - otherwise
    '*************************************************************
    
    Dim intNum As Long
    
    'Assume array isn't dimensioned.
    intNum = -1
    
    ' Avoid the run-time error if dynamic array isn't dimensioned.
    On Error Resume Next
    
    If IsEmpty(arr) Then
        IsArrayDimensioned = False
        Exit Function
    Else
        'Get the highest subscript.
        intNum = UBound(arr)
    End If
    
    'Reset error number
    On Error GoTo 0
    
    ' If the subscript is same as initial, the dynamic array isn't dimensioned.
    If intNum = -1 Then
        IsArrayDimensioned = False
    Else
        IsArrayDimensioned = True
    End If

End Function

Public Function IsElementInArray(ByRef arr As Variant, ByVal item As Variant) As Boolean

    ' Sprawdza, czy element jest w liœcie. Zwraca boolean
    ' arr - lista, która ma zostaæ sprawdzona
    ' varElement - element do sprawdzenia
    
    Dim varTemp As Variant
    
    IsElementInArray = False
    
    If IsArrayDimensioned(arr) Then
        For Each varTemp In arr
            If varTemp = item Then
                IsElementInArray = True
                Exit For
            End If
        Next varTemp
    End If

End Function

Public Function ArrayAddItem(ByVal arr As Variant, ByVal item As Variant) As Variant

    ' Dodawanie elementu do listy elementów. Zwraca listê elementów
    ' arr - lista, do której ma zostaæ dodany element
    ' varElement - element do dodania

    Dim lngMax As Long
    
    If IsArrayDimensioned(arr) = False Then
        ReDim arr(0)
        lngMax = 0
    Else
        lngMax = UBound(arr, 1) + 1
        ReDim Preserve arr(lngMax)
    End If
    
    arr(lngMax) = item
    
    ArrayAddItem = arr

End Function

Public Function ArrayRemoveDuplicates(ByRef arr As Variant) As Variant

    ' Usuwa duplikaty z listy. Argument varArray jest zmieniany w trakcie
    ' arr - lista do usuniêcia duplikatów

    Dim d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    
    Dim item As Variant
    For Each item In arr
        If Not d.Exists(item) Then
            d.Add item, 1
        End If
    Next item
    
    Dim arr2 As Variant
    ArrayRemoveDuplicates = d.Keys()

End Function

Public Function ArrayGetDuplicates(ByVal arr As Variant) As Variant

    ' Zwraca zduplikowane wartoœci z listy
    ' arr - lista, która ma zostaæ przeszukana pod k¹tem duplikatów
    
    Dim i As Long
    Dim j As Long
    Dim dups As Variant
    
    Set ArrayGetDuplicates = Nothing
    
    If IsArrayDimensioned(arr) = True Then
        For i = 0 To UBound(arr)
            If i + 1 <= UBound(arr) Then
                For j = (i + 1) To UBound(arr)
                    If arr(i) = arr(j) Then
                        dups = ArrayAddItem(dups, arr(i))
                    End If
                Next j
            End If
        Next i
        
        If IsArrayDimensioned(dups) = True Then ArrayGetDuplicates = dups
    End If

End Function

Public Function ArrayMaxValue(ByVal arr As Variant) As Variant

    ' Zwraca maksymaln¹ wartoœæ z tabeli
    ' arr - lista do przeszukania

    Dim i As Long
    Dim varMax As Variant
    
    If IsArrayDimensioned(arr) = True Then
        varMax = arr(0)
        For i = 0 To UBound(arr)
            If arr(i) > varMax Then varMax = arr(i)
        Next i
    End If
    
    ArrayMaxValue = varMax

End Function

Public Function ArrayMinValue(ByVal arr As Variant) As Variant

    ' Zwraca minimaln¹ wartoœæ z tabeli
    ' arr - lista do przeszukania
    
    Dim i As Long
    Dim varMin As Variant
    
    If IsArrayDimensioned(arr) = True Then
        varMin = arr(0)
        For i = 1 To UBound(arr)
            If arr(i) < varMin Then varMin = arr(i)
        Next i
    End If
    
    ArrayMinValue = varMin

End Function

Public Function ArraySort(ByVal arr As Variant, Optional ByVal ascending As Boolean = True) As Variant

    ' Sortuje elementy w podanej kolejnoœci.
    ' arr - lista do posortowania
    ' blnAscending - kolejnoœæ; True - rosn¹ca, False - malej¹ca

    Call QuicksortSub(arr, LBound(arr), UBound(arr))
    If Not (ascending) Then arr = ArrayReverseOrder(arr)
    ArraySort = arr

End Function

Private Function QuicksortSub(ByRef varArray As Variant, ByVal min As Long, ByVal max As Long) As Variant

    ' Sortowanie qsort
    ' varArray - lista do posortowania

    Dim med_value As String
    Dim hi As Long
    Dim lo As Long
    Dim i As Long
    
    ' If the list has only 1 item, it's sorted.
    If min >= max Then Exit Function
    
    ' Pick a dividing item randomly.
    i = min + Int(Rnd(max - min + 1))
    med_value = varArray(i)
    
    ' Swap the dividing item to the front of the list.
    varArray(i) = varArray(min)
    
    ' Separate the list into sublists.
    lo = min
    hi = max
    Do
        ' Look down from hi for a value < med_value.
        Do While varArray(hi) >= med_value
            hi = hi - 1
            If hi <= lo Then Exit Do
        Loop
        
        If hi <= lo Then
            ' The list is separated.
            varArray(lo) = med_value
            Exit Do
        End If
        
        ' Swap the lo and hi values.
        varArray(lo) = varArray(hi)
        
        ' Look up from lo for a value >= med_value.
        lo = lo + 1
        Do While varArray(lo) < med_value
            lo = lo + 1
            If lo >= hi Then Exit Do
        Loop
        
        If lo >= hi Then
            ' The list is separated.
            lo = hi
            varArray(hi) = med_value
            Exit Do
        End If
        
        ' Swap the lo and hi values.
        varArray(hi) = varArray(lo)
    Loop ' Loop until the list is separated.
    
    ' Recursively sort the sublists.
    Call QuicksortSub(varArray, min, lo - 1)
    Call QuicksortSub(varArray, lo + 1, max)

End Function

Public Function ArrayReverseOrder(ByVal varArray As Variant) As Variant

    ' Odwraca kolejnoœæ listy
    ' varArray - lista do odwrócenia
    
    Dim varTempArray As Variant
    Dim i As Long
    Dim lngMax As Long
    
    varTempArray = varArray
    
    lngMax = UBound(varArray)
    
    For i = LBound(varArray) To UBound(varArray)
        varTempArray(lngMax - i) = varArray(i)
    Next i
    
    ArrayReverseOrder = varTempArray

End Function

Public Function Array2DTranspose(avValues As Variant) As Variant

    ' Transpozycja macierzy 2 wymiarowej. Zwraca variant
    ' avValues - macierz do transponowania

    If IsArrayDimensioned(avValues) = False Then
        GoTo ErrFailed
    End If

    Dim lThisCol As Long, lThisRow As Long
    Dim lUb2 As Long, lLb2 As Long
    Dim lUb1 As Long, lLb1 As Long
    Dim avTransposed As Variant
    
    If IsArray(avValues) Then
        On Error GoTo ErrFailed
        lUb2 = UBound(avValues, 2)
        lLb2 = LBound(avValues, 2)
        lUb1 = UBound(avValues, 1)
        lLb1 = LBound(avValues, 1)
        
        ReDim avTransposed(lLb2 To lUb2, lLb1 To lUb1)
        For lThisCol = lLb1 To lUb1
            For lThisRow = lLb2 To lUb2
                avTransposed(lThisRow, lThisCol) = avValues(lThisCol, lThisRow)
            Next
        Next
    End If
    
    Array2DTranspose = avTransposed
    Exit Function

ErrFailed:
    Debug.Print err.Description
    ' Debug.Assert False
    
    Array2DTranspose = Empty
    Exit Function
    
    Resume
    
End Function

Public Function ArrayMerge(ByRef var1 As Variant, ByRef var2 As Variant) As Variant

    ' Scala dwie dwuwymiarowe macierze
    Dim var As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    
    If IsArrayDimensioned(var1) And IsArrayDimensioned(var2) Then
        If UBound(var1, 2) = UBound(var2, 2) Then
            i = UBound(var1, 1) + UBound(var2, 1) + 1
            j = UBound(var1, 2)
            
            ReDim var(i, j)
            
            For i = 0 To UBound(var1, 1)
                For j = 0 To UBound(var1, 2)
                    var(i, j) = var1(i, j)
                Next j
            Next i
            
            For k = 0 To UBound(var2, 1)
                For l = 0 To UBound(var2, 2)
                    var(i + k, l) = var2(k, l)
                Next l
            Next k
        End If
    End If
    
    ArrayMerge = var

End Function

Public Function ArrayTrimRows(ByVal var As Variant, ByVal intRowFrom As Integer) As Variant

    Dim var2 As Variant
    Dim i As Long
    Dim j As Long
    
    ReDim var2(UBound(var, 1) - intRowFrom, UBound(var, 2))
    intRowFrom = intRowFrom + LBound(var, 1)
    
    For i = LBound(var, 1) To UBound(var, 1)
        If i >= intRowFrom Then
            For j = LBound(var, 2) To UBound(var, 2)
                var2(i - intRowFrom, j) = var(i, j)
            Next j
        End If
    Next i
    
    ArrayTrimRows = var2

End Function

Public Function MultiDimToOneDimArray(ByRef arr As Variant) As Variant

    ' Konwertuje tablicê wielowymiarow¹ do tablicy jednowymiarowej

    Dim arr2 As Variant
    Dim i As Long
    Dim j As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            arr2 = Utils.ArrayAddItem(arr2, arr(i, j))
        Next j
    Next i
    MultiDimToOneDimArray = arr2

End Function

' =============================================================================
' Module Collections
' =============================================================================

Public Function CollectionGetKey(ByRef col As Collection, ByVal strKey As String, Optional ByVal defaultValue As Variant = Nothing) As Variant

    ' Zwraca wartosc dla podanego klucza kolekcji, ale tylko je¿eli kolekcja posiada
    ' taki klucz. W przeciwnym wypadku zwraca defaultValue
    ' col - kolekcja
    ' strKey - nazwa klucza
    
    If CollectionHasKey(col, strKey) Then
        If Not (IsNull(col(strKey))) Then
            Select Case VBA.VarType(col.item(strKey))
                Case vbObject:
                    Set CollectionGetKey = col(strKey)
                Case vbDataObject:
                    Set CollectionGetKey = col(strKey)
                Case Else
                    CollectionGetKey = col(strKey)
            End Select
            Exit Function
        End If
    End If
    
    CollectionGetKey = defaultValue

End Function

Public Function CollectionHasKey(ByRef col As Collection, ByVal strKey As String) As Boolean

    ' Sprawdza, czy podana kolekcja ma podany klucz
    ' col - kolekcja
    ' strKey - nazwa klucza
    
    Dim v As Variant
    
    On Error Resume Next
    
    v = col(strKey)
    
    If err.Number = 450 Or err.Number = 0 Then
        err.Number = 0
        err.Description = vbNullString
        err.Source = vbNullString
        CollectionHasKey = True
    Else
        CollectionHasKey = False
    End If

End Function

Public Function CollectionHasValue(ByRef col As Collection, ByVal value) As Boolean

    ' Sprawdza, czy podana kolekcja ma podan¹ wartoœæ
    ' col - kolekcja
    ' value - wartoœæ do sprawdzenia
    
    If Not col Is Nothing Then
        Dim i As Integer
        For i = 1 To col.count
            If isObject(value) Then
                If col(i) Is value Then
                    CollectionHasValue = True
                    Exit Function
                End If
            Else
                If col(i) = value Then
                    CollectionHasValue = True
                    Exit Function
                End If
            End If
        Next i
    End If
    
    CollectionHasValue = False

End Function

Public Function CollectionToArray(ByRef col As Collection) As Variant

    ' Zmienia kolekcjê na tablicê elementów
    Dim arr As Variant
    If Not col Is Nothing Then
        Dim item As Variant
        For Each item In col
            arr = Utils.ArrayAddItem(arr, item)
        Next item
        CollectionToArray = arr
    End If

End Function

Public Function CollectionToString(ByRef col As Collection, Optional ByVal delim As String) As String

    '...
    If Not IsNullOrEmpty(col) Then
        Dim text As String
        Dim item As Variant
        Dim i As Boolean
        For Each item In col
            If Not i Then
                text = CStr(item)
                i = True
            Else
                text = text & delim & CStr(item)
            End If
        Next item
        CollectionToString = text
    End If

End Function

Public Function VariantToCollection(ByRef var As Variant, Optional ByVal FirstRowAsColumnNames As Boolean) As Collection

    ' Zamienia macierz variant na kolekcjê s³owników
    ' var - macierz do zamiany
    ' blnFirstRowAsColumnNames - czy funkcja powinna przyj¹æ pierwszy wiersz jako
    ' nazwy kolumn; wtedy nazwy kolumn stanowi¹ klucze w s³owniku. Je¿eli fa³sz, wtedy
    ' przyjmuje numery kolumn jako klucze w s³owniku
    
    Dim col As New Collection
    Dim dict As Object
    Dim row As Long
    Dim column As Long
    
    If VarType(var) >= vbArray Then
        If Utils.IsArrayDimensioned(var) Then
            For row = LBound(var, 1) + IIf(FirstRowAsColumnNames, 1, 0) To UBound(var, 1)
                Set dict = CreateObject("Scripting.Dictionary")
                
                ' FIXME: czasami mog¹ byæ jednowymiarowe macierze
                For column = LBound(var, 2) To UBound(var, 2)
                    If FirstRowAsColumnNames Then
                        Call dict.Add(var(LBound(var, 1), column), var(row, column))
                    Else
                        Call dict.Add(column, var(row, column))
                    End If
                Next column
                
                Call col.Add(dict)
            Next row
        End If
    Else
        ' Zmienna nie jest macierz¹!
    End If
    
    Set VariantToCollection = col

End Function

'Public Function SetCollectionOfDictCompareMode(ByRef col As Collection, ByVal i As CompareMethod) As Collection
'
'    Dim d As Scripting.Dictionary
'    Dim d2 As Scripting.Dictionary
'    Dim col2 As Collection
'    Dim key As Variant
'
'    If Not col Is Nothing Then
'        Set col2 = New Collection
'
'        For Each d In col
'            Set d2 = New Scripting.Dictionary
'            d2.CompareMode = TextCompare
'            For Each key In d.Keys
'                d2.Add key, d(key)
'            Next key
'            col2.Add d2
'        Next d
'    End If
'
'    Set SetCollectionOfDictCompareMode = col2
'
'End Function

Public Function GetDictDictSafe(ByRef dict As Scripting.Dictionary, ByVal key As String) As Scripting.Dictionary
    If dict.Exists(key) Then
        If TypeName(dict(key)) = "Dictionary" Then
            Set GetDictDictSafe = dict(key)
            Exit Function
        End If
    End If
    Set GetDictDictSafe = New Scripting.Dictionary
End Function

Public Function GetDictDateSafe(ByRef dict As Scripting.Dictionary, ByVal key As String) As Variant
    Dim var As Variant
    If dict.Exists(key) Then
        Dim value As Variant
        value = dict(key)
        
        If Utils.IsNullOrEmpty(value) Then
            'pass
        ElseIf isObject(value) Then
            'pass
        Else
            var = ParseDate(value)
            var = VBA.DateSerial(VBA.Year(var), VBA.Month(var), VBA.Day(var))
        End If
    End If
    GetDictDateSafe = var
End Function

Public Function GetDictLongDateSafe(ByRef dict As Scripting.Dictionary, ByVal key As String) As Variant
    Dim var As Variant
    If dict.Exists(key) Then
        Dim value As Variant
        value = dict(key)

        If Utils.IsNullOrEmpty(value) Then
            'pass
        ElseIf isObject(value) Then
            'pass
        Else
            var = ParseDate(value)
        End If
    End If
    GetDictLongDateSafe = var
End Function

Public Function GetDictVariantSafe(ByRef dict As Scripting.Dictionary, ByVal key As String) As Variant
    Dim var As Variant
    If dict.Exists(key) Then
        If Utils.IsNullOrEmpty(dict(key)) Then
            'pass
        ElseIf isObject(dict(key)) Then
            'pass
        Else
            var = dict(key)
        End If
    End If
    GetDictVariantSafe = var
End Function

Public Function GetDictNumericSafe(ByRef dict As Scripting.Dictionary, ByVal key As String) As Variant
    Dim var As Variant
    If dict.Exists(key) Then
        If Utils.IsNullOrEmpty(dict(key)) Then
            'pass
        ElseIf isObject(dict(key)) Then
            'pass
        Else
            var = dict(key)
            If Not IsNumeric(var) Then
                Dim delim As String
                delim = GetDecimalDelim
                
                With New RegExp
                    .IgnoreCase = True
                    .MultiLine = True
                    .pattern = "[^0-9|\" & delim & "|\.]"
                    var = .Replace(var, "")
                    var = Replace(var, ".", delim)
                End With
            End If
        End If
    End If
    GetDictNumericSafe = var
End Function

Public Function GetSubdictionary(ByRef dict As Scripting.Dictionary, ByVal prefix As String) As Scripting.Dictionary

    Dim d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    
    Dim key As Variant
    Dim key2 As String
    For Each key In dict.Keys
        If key Like prefix & "*" Then
            key2 = Right(key, Len(key) - Len(prefix))
            'key2 = Replace(key, prefix, "")
            d.Add key2, dict(key)
        End If
    Next key
    Set GetSubdictionary = d

End Function

' =============================================================================
' Module DateTime
' =============================================================================

Public Function ParseDate(ByVal dateValue As String) As Variant

    On Error Resume Next
    If IsDate(dateValue) Then
        ParseDate = CDate(dateValue)
        Exit Function
        
    Else
        On Error GoTo NOT_ISO
        Dim temp As Date
        temp = JsonConverter.ParseIso(dateValue)
        Exit Function
        
NOT_ISO:
        ' np. 2021-02-18 22:35:04.000
        If Right(dateValue, 4) = ".000" Then
            ParseDate = CDate(Left(dateValue, Len(dateValue) - 4))
            Exit Function
        End If
        
    End If

End Function

' =============================================================================
' Others
' =============================================================================

Public Function GetRandomString(ByVal intLength As Integer) As String
    Randomize
    Dim seed As String
    seed = "abcdefghijklmnopqrstuvwxyz"
    'seed = seed & UCase(seed) & "0123456789"
    
    Dim i As Long
    For i = 1 To intLength
        GetRandomString = GetRandomString & Mid$(seed, Int(Rnd() * Len(seed) + 1), 1)
    Next
End Function

Public Function GetDecimalDelim() As String
    GetDecimalDelim = Mid$(CStr(1 / 2), 2, 1)
End Function

Public Function GetDateFormat() As String
    ' Zwraca domyslny format daty
    Dim temp As Date
    temp = DateSerial(2000, 11, 12)
    Dim f As String
    f = CStr(temp)
    f = Replace(f, "2000", "YYYY")
    f = Replace(f, "11", "MM")
    f = Replace(f, "12", "DD")
    GetDateFormat = f
End Function

Public Function GetNumberFormat(Optional ByVal numOfDigits As Integer = 0) As String
    ' Zwraca domyslny format liczby
    Dim num As Double
    num = 1 / 3
    
    Dim f As String
    f = CStr(num)
    
    If numOfDigits <> 0 Then
        f = Left(num, 2 + numOfDigits)
    Else
        f = Left(num, 1)
    End If
    f = Replace(f, "3", "0")
    GetNumberFormat = f
End Function

Public Function GetLastColumn(ByVal Wks As Worksheet, ByVal StartRow As Long, ByVal StartColumn As Long) As Long
    Dim cell As Range
    Set cell = Wks.Cells(StartRow, StartColumn)
    Do While Len(cell.value) <> 0
        Set cell = cell.Offset(0, 1)
    Loop
    GetLastColumn = cell.column - 1
End Function

Public Function GetLastRow(ByVal Wks As Worksheet, ByVal StartRow As Long, ByVal StartColumn As Long, ByVal EndColumn As Long) As Long
    Dim cell As Range
    Set cell = Wks.Range(Wks.Cells(StartRow, StartColumn), Wks.Cells(StartRow, EndColumn))
    Do While cell.Cells.count - Application.WorksheetFunction.CountBlank(cell) <> 0
        Set cell = cell.Offset(1, 0)
    Loop
    GetLastRow = cell.row - 1
End Function

Public Function IsNullOrEmpty(ByRef value As Variant) As Boolean
    ' Sprawdza, czy podana wartosc jest pusta
    
    Select Case VBA.VarType(value)
    Case 0
        IsNullOrEmpty = True
        
    Case VBA.vbEmpty
        IsNullOrEmpty = True
        
    Case VBA.vbNull
        IsNullOrEmpty = True
        
    Case VBA.vbDate
        IsNullOrEmpty = False
        
    Case VBA.vbString
        IsNullOrEmpty = (Len(value) = 0)
        
    Case VBA.vbBoolean
        IsNullOrEmpty = False
        
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        IsNullOrEmpty = Not Utils.IsArrayDimensioned(value)
        
    Case VBA.vbVariant
        IsNullOrEmpty = IsEmpty(value)
        
    Case VBA.vbObject
        IsNullOrEmpty = value Is Nothing
        
    Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal, VBA.vbByte
        IsNullOrEmpty = False
        
    Case Else
        ' vbError, vbDataObject, vbUserDefinedType
        ' Use VBA's built-in to-string
        On Error Resume Next
        IsNullOrEmpty = value Is Nothing
        On Error GoTo 0
        
    End Select
    
End Function

Public Function PrintF(ByVal strText As String, ParamArray args()) As String
    ' © codekabinett.com - You may use, modify, copy, distribute this code as long as this line remains
    Dim i           As Integer
    Dim strRetVal   As String
    Dim startPos    As Integer
    Dim endPos      As Integer
    Dim formatString As String
    Dim argValueLen As Integer
    strRetVal = strText
    
    Dim myArgs As Variant
    If IsArray(args(0)) Then
           myArgs = args(0)
       Else
           myArgs = args()
    End If
    
    For i = LBound(myArgs) To UBound(myArgs)
        argValueLen = Len(CStr(i))
        startPos = InStr(strRetVal, "{" & CStr(i) & ":")
        If startPos > 0 Then
            endPos = InStr(startPos + 1, strRetVal, "}")
            formatString = Mid(strRetVal, startPos + 2 + argValueLen, endPos - (startPos + 2 + argValueLen))
            strRetVal = Mid(strRetVal, 1, startPos - 1) & Format(myArgs(i), formatString) & Mid(strRetVal, endPos + 1)
        Else
            strRetVal = Replace(strRetVal, "{" & CStr(i) & "}", myArgs(i))
        End If
    Next i

    PrintF = strRetVal
End Function

Public Function TryParseJson(ByVal json As String, ByRef obj As Object) As Boolean
    On Error GoTo ErrHandler
    Set obj = ParseJson(json)
    TryParseJson = True
    Exit Function
ErrHandler:
    err.Clear
End Function

Public Sub Wait(ByVal miliseconds As Integer)
    Call Sleep(miliseconds)
End Sub

Public Sub RandomSleep(ByVal secondsMin As Integer, ByVal secondsMax As Integer)
    Dim seconds As Integer
    VBA.Randomize
    seconds = Int((secondsMax - secondsMin + 1) * Rnd + secondsMin)
    Call Sleep(seconds)
End Sub

Public Function CreateGUID() As String
    Do While Len(CreateGUID) < 32
        If Len(CreateGUID) = 16 Then
            '17th character holds version information
            CreateGUID = CreateGUID & hex$(8 + CInt(Rnd * 3))
        End If
        CreateGUID = CreateGUID & hex$(CInt(Rnd * 15))
    Loop
    CreateGUID = "{" & Mid(CreateGUID, 1, 8) & "-" & Mid(CreateGUID, 9, 4) & "-" & Mid(CreateGUID, 13, 4) & "-" & Mid(CreateGUID, 17, 4) & "-" & Mid(CreateGUID, 21, 12) & "}"
End Function

'---------------------------------------------------------------------------------------
' Procedure : CountOccurrences
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Count the number of times a string is found within a string
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
'             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
' Req'd Refs: None required
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sText         String to search through
' sSearchTerm   String to count the number of occurences of
'
' Usage:
' ~~~~~~
' CountOccurrences("aaa", "a")               -> 3
' CountOccurrences("514-55-55-5555-5", "-")  -> 4
' CountOccurrences("192.168.2.1", ".")       -> 3
' CountOccurrences("192.168.2.1", "/")       -> 0
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2016-02-19              Initial Release
' 2         2019-02-16              Updated Copyright
'                                   Updated Error Handler
'---------------------------------------------------------------------------------------
Public Function CountOccurrences(sText As String, sSearchTerm As String) As Long
    On Error GoTo Error_Handler
 
    CountOccurrences = UBound(Split(sText, sSearchTerm))
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: CountOccurrences" & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Public Sub ToClipboard(text As String)
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub


Public Function IdComp(ByRef value1 As Object, ByRef value2 As Object) As Boolean
    ' Porownuje dwa obiekty wedlug pola Id
    If Not value1 Is Nothing And Not value2 Is Nothing Then
        IdComp = (value1.Id = value2.Id)
    End If
End Function

Public Function VarComp(ByVal value1 As Variant, ByVal value2 As Variant) As Boolean
    ' Bezpieczne porownywanie dwoch zmiennych typu variant, ktore nie sa tablicami
    If IsArray(value1) Or IsArray(value2) Then
        VarComp = (value1 = value2)
    Else
        If IsEmpty(value1) And Not IsEmpty(value2) Then
            VarComp = False
        ElseIf IsEmpty(value2) And Not IsEmpty(value1) Then
            VarComp = False
        Else
            If IsNumeric(value1) And IsNumeric(value2) Then
                VarComp = (CDbl(value1) = CDbl(value2))
                Exit Function
            End If
            
            If IsDate(value1) And IsDate(value2) Then
                VarComp = (CDate(value1) = CDate(value2))
                Exit Function
            End If
            
            VarComp = (CStr(value1) = CStr(value2))
        End If
    End If
End Function
