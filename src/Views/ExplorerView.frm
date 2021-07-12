VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExplorerView 
   Caption         =   "UserForm1"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "ExplorerView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExplorerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Views")
Option Explicit

Implements IExplorerView

Private Type TFields
    Model As IExplorerViewModel
    IsCancelled As Boolean
    selectMode As ESelectMode
End Type
Private this As TFields

Public Property Get Model() As IExplorerViewModel
    Set Model = this.Model
End Property
Private Property Get IExplorerView_Model() As IExplorerViewModel
    Set IExplorerView_Model = Model
End Property
Private Property Let Model(ByVal value As IExplorerViewModel)
    Set this.Model = value
End Property

Public Property Get IsCancelled() As Boolean
    IsCancelled = False
End Property
Private Property Get IExplorerView_IsCancelled() As Boolean
    IExplorerView_IsCancelled = IsCancelled
End Property
Private Property Let IsCancelled(ByVal value As Boolean)
    this.IsCancelled = value
End Property

Public Property Get Self() As ExplorerView
    Set Self = Me
End Property
Private Property Get IExplorerView_Self() As IExplorerView
    Set IExplorerView_Self = Self
End Property

Public Sub Init(ByRef cModel As IExplorerViewModel, _
                ByVal title As String, _
                ByVal multiselect As Boolean, _
                Optional ByVal cSelectMode As ESelectMode = ESelectMode.ESelectModeAll)
                
    GuardClauses.IsNothing cModel, TypeName(cModel)

    Model = cModel
    Me.Caption = title
    If multiselect Then
        ListBox.multiselect = fmMultiSelectExtended
    Else
         ListBox.multiselect = fmMultiSelectSingle
    End If
    this.selectMode = cSelectMode
End Sub

Public Sub Display()
    Me.RefreshButton.Caption = ChrW(&HE72C)
    RefreshView
    Me.Show
End Sub
Private Sub IExplorerView_Display()
    Display
End Sub

Public Sub HideView()
    Me.Hide
End Sub
Private Sub IExplorerView_HideView()
    HideView
End Sub

Public Sub OK()
    IsCancelled = False
    Dim col As Collection
    Set col = GetSelectedItems
    Set col = FilterSelectedItems(col, this.selectMode)
    Model.SetSelectedItems col
    HideView
End Sub

Public Sub Cancel()
    IsCancelled = True
    HideView
End Sub

Private Sub UpdateView()
    Dim data As Variant
    data = IDriveItemCollectionToVariantArray
    ListBox.List = data
    ListBox.ColumnCount = UBound(data, 1) + 1
    ListBox.ColumnHeads = False
    
    Dim widths As Variant
    widths = GetListboxColumnsWidth(data)
    widths(0) = 0 ' hide ID column
    ListBox.ColumnWidths = Join(widths, ";")
    PathTextBox.text = Model.CurrentItem.path
End Sub

Private Sub RefreshView()
    Dim fldr As IFolder
    Set fldr = Model.CurrentItem
    
    Dim col As Collection
    Set col = fldr.GetChildren
    
    Model.SetItems col
    UpdateView
End Sub

Private Sub ChangeFolder()
    Dim selected As Collection
    Set selected = GetSelectedItems
    
    If Not selected Is Nothing Then
        If selected.Count <> 0 Then
            Dim item As IDriveItem
            Set item = selected(1)
            
            If item.IsFolder Then
                Model.SetCurrentItem item
                RefreshView
            End If
        End If
    End If
End Sub

' Event handlers
Private Sub OkButton_Click()
    OK
End Sub

Private Sub RefreshButton_Click()
    RefreshView
End Sub

Private Sub ListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ChangeFolder
End Sub

Private Sub UserForm_QueryClose(cCancel As Integer, CloseMode As Integer)
    If Not IsCancelled Then
        cCancel = True
        Cancel
    End If
End Sub

Private Sub UserForm_Resize()
    ' TODO
End Sub

' Helpers
Private Function GetSelectedItems() As Collection
    Dim col As Collection
    Set col = New Collection
    
    Dim id As String
    Dim item As IDriveItem
    
    Dim i As Long
    For i = 1 To UBound(ListBox.List, 1)
        If ListBox.selected(i) Then
            id = ListBox.List(i, 0)
            Set item = GetItemFromId(id)
            col.Add item
        End If
    Next i
    
    Set GetSelectedItems = col
End Function

Private Function FilterSelectedItems(ByRef col As Collection, ByVal mode As ESelectMode) As Collection

    Dim col2 As Collection
    
    If mode = ESelectMode.ESelectModeAll Then
        Set col2 = col
    
    Else
        Dim item As IDriveItem
        For Each item In col
            Select Case mode
            Case ESelectMode.ESelectModeFilesOnly
                If item.IsFile Then col2.Add item
                
            Case ESelectMode.ESelectModeFoldersOnly
                If item.IsFolder Then col2.Add item
                
            End Select
        Next item
        
    End If
    
    Set FilterSelectedItems = col2

End Function

Private Function GetItemFromId(ByVal id As String) As IDriveItem
    If Not Model.CurrentItem.parent Is Nothing Then
        If id = Model.CurrentItem.parent.id Then
            Set GetItemFromId = Model.CurrentItem.parent
            Exit Function
        End If
    End If
    
    Dim item As IDriveItem
    For Each item In Model.items
        If item.id = id Then
            Set GetItemFromId = item
            Exit Function
        End If
    Next item
End Function

Private Function IDriveItemCollectionToVariantArray() As Variant
    Dim arr As Variant
    If Not Model.items Is Nothing Then
        Dim arrItemsCount As Long
        arrItemsCount = Model.items.Count
        
        ReDim arr(arrItemsCount, 3)

        Dim i As Long
        i = 1
        If Not Model.CurrentItem.parent Is Nothing Then
            ReDim arr(arrItemsCount + 1, 3)
            arr(1, 0) = Model.CurrentItem.parent.id
            arr(1, 1) = ".."
            i = 2
        End If
        
        arr(0, 0) = "id"
        arr(0, 1) = "Name"
        arr(0, 2) = "Size"
        arr(0, 3) = "Modification time"
        
        Dim fld As IFolder
        Dim fle As IFile
        Dim item As IDriveItem
        For Each item In Model.items
            If item.IsFile Then
                Set fle = item
                arr(i, 0) = fle.id
                arr(i, 1) = fle.name
                arr(i, 2) = fle.Size \ 1024
                arr(i, 3) = fle.LastModifiedTime
            Else
                Set fld = item
                arr(i, 0) = fld.id
                arr(i, 1) = fld.name
                arr(i, 2) = ""
                arr(i, 3) = ""
            End If
            
            i = i + 1
        Next item
        
    Else
        ReDim arr(0, 3)
        arr(0, 0) = "id"
        arr(0, 1) = "Name"
        arr(0, 2) = "Size"
        arr(0, 3) = "Modification time"
        
    End If
    
    IDriveItemCollectionToVariantArray = arr
End Function

Private Function GetListboxColumnsWidth(ByVal data As Variant) As Variant

    On Error Resume Next
       
    Dim widths As Variant
    Dim max As Integer
    Dim i As Integer
    Dim j As Integer
    For i = LBound(data, 2) To UBound(data, 2)
        For j = LBound(data, 1) To UBound(data, 1)
            If max < Len(data(j, i)) Then max = Len(data(j, i))
        Next j
                    
        If max >= 0 And max < 5 Then
            widths = Utils.ArrayAddItem(widths, 30)
            
        ElseIf max >= 5 And max < 10 Then
            widths = Utils.ArrayAddItem(widths, 60)
            
        ElseIf max >= 10 And max < 20 Then
            widths = Utils.ArrayAddItem(widths, 100)
            
        ElseIf max >= 20 And max < 30 Then
            widths = Utils.ArrayAddItem(widths, 130)
            
        ElseIf max >= 30 And max < 50 Then
            widths = Utils.ArrayAddItem(widths, 160)
            
        ElseIf max >= 50 And max < 100 Then
            widths = Utils.ArrayAddItem(widths, 190)
            
        ElseIf max >= 100 Then
            widths = Utils.ArrayAddItem(widths, 220)
            
        End If
        
        max = 0
    Next i
    
    GetListboxColumnsWidth = widths

End Function
