Attribute VB_Name = "mListBoxTools"
Option Explicit


Public Sub LBXSelectItems(lst As MSForms.ListBox, arr As Variant, Optional colNDX As Integer = 0)
    Dim x As Long, y As Long
    With lst
        For x = 0 To UBound(arr)
            For y = 0 To .ListCount - 1
                If .List(y, colNDX) = arr(x) Then
                    .Selected(y) = True
                    Exit For
                End If
            Next y
        Next x
    End With
End Sub

Public Function LBXSelectCount(LBX As MSForms.ListBox) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LBXSelectCount
' Returns the number of selected items in LBX.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim N As Long
    Dim C As Long
    With LBX
        For N = 0 To .ListCount - 1
            If .Selected(N) = True Then
                C = C + 1
            End If
        Next N
    End With
    LBXSelectCount = C
End Function

Public Function LBXSelectedItems(LBX As MSForms.ListBox, Optional colNDX As Integer = 0) As String()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SelectedItems
' This returns a 0-based array of strings, each of which is a selected
' item in the list box. If LBX is empty, the result is an unallocated
' array. The caller should first call SelectionInfo to determine whether
' there are any selected items prior to calling this procedure.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim SelCount As Long
    Dim FirstIndex As Long
    Dim LastIndex As Long
    Dim SelItems() As String
    Dim Ndx As Long
    Dim ArrNdx As Long
    
    ''''''''''''''''''''''''
    ' If list is empty, get
    ' out now.
    ''''''''''''''''''''''''
    If LBX.ListCount = 0 Then
        Exit Function
    End If
    
'    If LBX.ColumnCount > 1 Then
'        ''''''''''''''''''''''''''
'        ' No support for mutliple
'        ' column listboxes.
'        ''''''''''''''''''''''''''
'        Exit Function
'    End If
    
    LBXSelectionInfo LBX:=LBX, SelectedCount:=SelCount, _
        FirstSelectedItemIndex:=FirstIndex, LastSelectedItemIndex:=LastIndex
    
    ''''''''''''''''''''''''''''''''''''
    ' If nothing was selected, get out.
    ''''''''''''''''''''''''''''''''''''
    If SelCount = 0 Then
        Exit Function
    End If
    
    
    ArrNdx = 0
    '''''''''''''''''''''''''''''''''''
    ' Redim the result array to the
    ' number of selected items. This
    ' array is 0-based.
    '''''''''''''''''''''''''''''''''''
    ReDim SelItems(0 To SelCount - 1)
    
    With LBX
        For Ndx = 0 To .ListCount - 1
            If .Selected(Ndx) = True Then
                SelItems(ArrNdx) = SetEmptyValue(.List(Ndx, colNDX), NullStrings)
                ArrNdx = ArrNdx + 1
            End If
        Next Ndx
    End With
    
    LBXSelectedItems = SelItems

End Function

Public Sub LBXSelectionInfo(LBX As MSForms.ListBox, ByRef SelectedCount As Long, _
    ByRef FirstSelectedItemIndex As Long, ByRef LastSelectedItemIndex As Long)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SelectionInfo
' This procedure provides information about the selected
' items in the listbox referenced by LBX. The variable
' SelectedCount will be populated with the number of selected
' items, the variable FirstSelectedItem will be popuplated
' with the index number of the first (from the top down)
' selected item, and the variable LastSelectedItem will return
' the index number of the last (from the top down) selected
' item. If no item(s) are selected or ListIndex < 0,
' SelectedCount is set to 0, and FirstSelectedItem and
' LastSelectedItem are set to -1.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim FirstItem As Long: FirstItem = -1
Dim LastItem As Long:   LastItem = -1
Dim SelCount As Long:   SelCount = 0
Dim Ndx As Long

''''''''''''''''''''''''
' If list is empty, get
' out now.
''''''''''''''''''''''''
If LBX.ListCount = 0 Then
    Exit Sub
End If

With LBX
    If .ListCount = 0 Then
        SelectedCount = 0
        FirstSelectedItemIndex = -1
        LastSelectedItemIndex = -1
        Exit Sub
    End If
    If .ListIndex < 0 Then
        SelectedCount = 0
        FirstSelectedItemIndex = -1
        LastSelectedItemIndex = -1
        Exit Sub
    End If
    For Ndx = 0 To .ListCount - 1
        If .Selected(Ndx) = True Then
            If FirstItem < 0 Then
                FirstItem = Ndx
            End If
            SelCount = SelCount + 1
            LastItem = Ndx
        End If
    Next Ndx
End With
    
SelectedCount = SelCount
FirstSelectedItemIndex = FirstItem
LastSelectedItemIndex = LastItem

End Sub

Public Sub LBXSelectAllItems(LBX As MSForms.ListBox)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SelectAllItems
' This procedure selects all items in the listbox LBX.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Long
With LBX
    ''''''''''''''''''''''''
    ' If list is empty, get
    ' out now.
    ''''''''''''''''''''''''
    If .ListCount = 0 Then
        Exit Sub
    End If
    For Ndx = 0 To .ListCount - 1
        .Selected(Ndx) = True
    Next Ndx
End With

End Sub

Public Function LBXSelectedIndexes(LBX As MSForms.ListBox) As Long()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LBXSelectedIndexes
' This returns an array of Longs that are the index numbers of
' the selected items.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim L() As Long
Dim C As Long
Dim N As Long
If LBXSelectCount(LBX) = 0 Then
    LBXSelectedIndexes = L
    Exit Function
End If
With LBX
    ReDim L(0 To .ListCount - 1)
    C = -1
    For N = 0 To .ListCount - 1
        If .Selected(N) Then
            C = C + 1
            L(C) = N
        End If
    Next N
End With
ReDim Preserve L(0 To C)
LBXSelectedIndexes = L

End Function

Public Sub LBXInvertSelection(LBX As MSForms.ListBox)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LBXInvertSelection
' Inverts selected items. Selected items are unselected and selected
' items are unselected.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
With LBX
    ''''''''''''''''''''''''
    ' If list is empty, get
    ' out now.
    ''''''''''''''''''''''''
    If .ListCount = 0 Then
        Exit Sub
    End If
    For Ndx = 0 To .ListCount - 1
        .Selected(Ndx) = Not .Selected(Ndx)
    Next Ndx
End With

End Sub

Public Function GetArrayFromList(ctrl As MSForms.ListBox, Optional intCol As Integer = 0) As Variant
    Dim arr() As Variant
    Dim i As Integer
    Dim varItem As Variant
    
    If ctrl.ItemsSelected.Count <> 0 Then
        i = 0
        For Each varItem In ctrl.ItemsSelected
            ReDim Preserve arr(i)
            arr(i) = ctrl.Column(intCol, varItem)
            i = i + 1
        Next varItem
    End If
    
    GetArrayFromList = arr
End Function
