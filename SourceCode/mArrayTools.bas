Attribute VB_Name = "mArrayTools"
Option Explicit

Public Function ConvertSQLArrToListArr(arr As Variant) As Variant
    Dim x As Long, y As Long
    Dim arrTemp As Variant
    
    ReDim arrTemp(UBound(arr, 2), UBound(arr))
    For x = 0 To UBound(arr, 2)
        For y = 0 To UBound(arr)
            arrTemp(x, y) = arr(y, x)
        Next y
    Next x
    
    ConvertSQLArrToListArr = arrTemp
End Function

' Check if item is in array
Public Function IsInArray(arr As Variant, item As Variant) As Boolean
    Dim i As Integer
    
    IsInArray = False
    
    For i = 0 To UBound(arr)
        If arr(i) = item Then
            IsInArray = True
            Exit For
        End If
    Next i
End Function

Public Sub PopulateListFromArray(arr As Variant, ctrl As Control)
    Dim x As Long, y As Long
    
    If IsArrayAllocated(arr) Then
'        ' Add Select all
'        ctrl.AddItem
'        ctrl.List(0, 0) = "(Select All)"
        
        For x = 0 To UBound(arr)
            ctrl.AddItem
            For y = 0 To UBound(arr, 2)
                ctrl.List(x, y) = IIf(IsNull(arr(x, y)), vbNullString, arr(x, y))
            Next y
            'ctrl.Selected(x + 1) = True
        Next x
        
'        ' Select all
'        Call LBXSelectAllItems(ctrl)
    End If
End Sub

Public Function GetArrayList(qry As String, Optional blnMultiColumn As Boolean = False, _
                                            Optional blnConnLocal As Boolean = True) As Variant
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim arr1 As Variant
Dim arr2 As Variant
Dim x As Long, y As Long

If gEnableErrorHandling Then On Error GoTo Err_Handler
PushCallStack "mArrayTools|GetArrayList"

If blnConnLocal = True Then
    Set conn = cn
Else
    Set conn = cnRemote
End If

Set rs = New ADODB.Recordset
With rs

    .Open qry, ActiveConnection:=conn
    If Not .EOF Then
        arr1 = .GetRows(100000)
        
        Select Case blnMultiColumn
            Case False
                ReDim arr2(UBound(arr1, 2))
                For x = 0 To UBound(arr1, 2)
                    arr2(x) = arr1(0, x)
                Next x
            Case True
                ReDim arr2(UBound(arr1, 2), UBound(arr1))
                For x = 0 To UBound(arr1)
                    For y = 0 To UBound(arr1, 2)
                        arr2(y, x) = arr1(x, y)
                    Next y
                Next x
        End Select
        
        GetArrayList = arr2
    End If
    .Close
End With

Set rs = Nothing

Proc_Exit:
PopCallStack
Exit Function

Err_Handler:
GlobalErrHandler
Resume Proc_Exit
End Function

Public Function IsArrayAllocated(arr As Variant) As Boolean
    Dim N As Long
    On Error Resume Next
    
    ' if Arr is not an array, return FALSE and get out.
    If IsArray(arr) = False Then
        IsArrayAllocated = False
        Exit Function
    End If
    
    ' Attempt to get the UBound of the array. If the array has not been allocated,
    ' an error will occur. Test Err.Number to see if an error occurred.
    N = UBound(arr, 1)
    If (Err.Number = 0) Then
        ''''''''''''''''''''''''''''''''''''''
        ' Under some circumstances, if an array
        ' is not allocated, Err.Number will be
        ' 0. To acccomodate this case, we test
        ' whether LBound <= Ubound. If this
        ' is True, the array is allocated. Otherwise,
        ' the array is not allocated.
        '''''''''''''''''''''''''''''''''''''''
        If LBound(arr) <= UBound(arr) Then
            ' no error. array has been allocated.
            IsArrayAllocated = True
        Else
            IsArrayAllocated = False
        End If
    Else
        ' error. unallocated array
        IsArrayAllocated = False
    End If

End Function
