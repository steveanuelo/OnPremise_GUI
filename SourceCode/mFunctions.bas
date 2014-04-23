Attribute VB_Name = "mFunctions"
Option Explicit

Global blnCalledFromUpdateFilter As Boolean

Const CBO_EXTENDED_WIDTH = 150

Public Enum enumNulls
    Uninitialized = 0
    NoValidData = 1
    NullStrings = 2
    NullCharacter = 3
    NullForDB = 4
    ZeroValue = 5
End Enum

' Set listbox extension event handler
Public Sub InitializeListBoxExtension(colControls As MSForms.Controls, lst As MSForms.ListBox, _
                                      dblExtendedHeight As Double, dblExtendedWidth As Double, _
                                      col As Collection)
    Dim strTypeName As String
    Dim objControl As MSForms.Control
    Dim clsEvents As C_ExtendListBoxDimensions
    
    If col Is Nothing Then
        Set col = New Collection
    End If
    
    'Loop through all the controls
    For Each objControl In colControls
        ' Exclude handle in the target list box
        If objControl.Name <> lst.Name Then
            strTypeName = TypeName(objControl)
            Select Case strTypeName
                Case "Label", "ComboBox", "TextBox", "ListBox", "CommandButton", "Frame"
                    'Create a new instance of the event handler class
                    Set clsEvents = New C_ExtendListBoxDimensions
                    
                    ' Set target listbox
                    Set clsEvents.MainListBox = lst
                    clsEvents.OriginalHeight = lst.Height
                    clsEvents.ExtendedHeight = dblExtendedHeight
                    clsEvents.OriginalWidth = lst.Width
                    clsEvents.ExtendedWidth = dblExtendedWidth
                    
                    'assign the right property for the event handler
                    Select Case strTypeName
                        Case "Label"
                            Set clsEvents.Label = objControl
                        Case "TextBox"
                            Set clsEvents.TextBox = objControl
                        Case "ListBox"
                            Set clsEvents.ListBox = objControl
                        Case "ComboBox"
                            Set clsEvents.ComboBox = objControl
                        Case "CommandButton"
                            Set clsEvents.CommandButton = objControl
                        Case "Frame"
                            Set clsEvents.Frame = objControl
                        Case Else
                            'Do nothing
                    End Select
                    
                    'Add the event handler instance to our collection,
                    'so it stays alive during the life of the workbook
                    col.Add clsEvents
                Case Else
                    'Do nothing
            End Select
        End If
    Next objControl
End Sub

Public Sub ClearList(ctrl As ComboBox)
    'On Error Resume Next
    Dim i As Integer

    For i = 1 To ctrl.ListCount
        ctrl.RemoveItem 0
    Next i
End Sub

Public Function GetPromoDate(dte As enumPromoDate, frm As frmMain) As String

    GetPromoDate = "1-Jan-1900"
    
    Select Case dte
        Case enumPromoDate.Start_Date
            GetPromoDate = frm.txtFromDate
            If Len(frm.txtFromExtnDate) <> 0 Then GetPromoDate = frm.txtFromExtnDate
        Case enumPromoDate.End_Date
            GetPromoDate = frm.txtToDate
            If Len(frm.txtToExtnDate) <> 0 Then GetPromoDate = frm.txtToExtnDate
    End Select

End Function

Public Function GetIN_List(arr As Variant, strTextQualifier As String, Optional strDelimiter As String = ",") As String
    Dim strIN As String
    Dim i As Integer
    
    strIN = vbNullString
    For i = 0 To UBound(arr)
        strIN = strIN & strTextQualifier & arr(i) & strTextQualifier
        If i < UBound(arr) Then
            strIN = strIN & strDelimiter
        End If
    Next i
        
    GetIN_List = strIN
End Function

Public Function GetLastRow(rng As Range) As String
    Dim z As Range
     
    Set z = rng.EntireColumn.Find("*", SearchDirection:=xlPrevious)
    
    If Not z Is Nothing Then
        GetLastRow = z.Address(False, False)
    Else
        GetLastRow = "No Data"
    End If
End Function

Public Function AddSelectAll(tbl As String, intCols As Integer) As String
    Dim i As Integer
    Dim nextCol As String
    
    nextCol = vbNullString
    For i = 1 To intCols - 1
        nextCol = nextCol & ", ''"
    Next i
    
    AddSelectAll = "SELECT DISTINCT '(Select All)'" & nextCol & " FROM " & tbl & " UNION ALL "
End Function

Public Function AddDummyOutlet(tbl As String) As String
    AddDummyOutlet = "SELECT DISTINCT 'Opportunity Outlet', '000000', '0000000000'" & " FROM " & tbl & " UNION ALL "
End Function

Public Sub ExtendWidth(ctrl As MSForms.Control, intOrigWidth As Integer)
    With ctrl
        Select Case .Width
            Case intOrigWidth
                .Width = CBO_EXTENDED_WIDTH
            Case CBO_EXTENDED_WIDTH
                .Width = intOrigWidth
        End Select
    End With
End Sub


' Returns an array of unique values within a specified range
Public Function uniqueValues(InputRange As Range)
    Dim cell As Range
    Dim tempList As Variant: tempList = ""
    For Each cell In InputRange
        If cell.Value <> "" Then
            If InStr(1, tempList, cell.Value) = 0 Then
                If tempList = "" Then tempList = Trim(CStr(cell.Value)) Else tempList = tempList & "|" & Trim(CStr(cell.Value))
            End If
        End If
    Next cell
    uniqueValues = Split(tempList, "|")
End Function

Public Function SetEmptyValue(val As Variant, Optional valueIfNull As enumNulls = 2) As Variant
    Dim strValIfNull As String
    Select Case valueIfNull
        Case Uninitialized
            strValIfNull = vbEmpty
        Case NoValidData
            strValIfNull = vbNull
        Case NullStrings
            strValIfNull = vbNullString
        Case NullCharacter
            strValIfNull = vbNullChar
        Case NullForDB
            strValIfNull = "NULL"
        Case ZeroValue
            strValIfNull = 0
    End Select
    
    SetEmptyValue = IIf(Len(val) <> 0, val, strValIfNull)
End Function

Public Function RoundNum(varNum As Variant, lngDecimalPlaces As Long) As Variant
    Dim dblResult As Double
    dblResult = Round(SetEmptyValue(varNum, ZeroValue), lngDecimalPlaces)
    
    If dblResult = 0 Then
        RoundNum = vbNullString
    Else
        RoundNum = dblResult
    End If
End Function

Public Function ConvToDbl(varNum As Variant)
    ConvToDbl = CDbl(SetEmptyValue(varNum, ZeroValue))
End Function

Public Function GetTagInfo(strTag As String, index As Integer, Optional Delimiter As String = "|") As String
    Dim arr As Variant
    arr = Split(strTag, Delimiter)
    If IsArrayAllocated(arr) Then
        GetTagInfo = arr(index)
    End If
End Function

Public Function GetRangeAddress(rng As Range, Optional blnAbsoluteAddress As Boolean = False) As String
If gEnableErrorHandling Then On Error GoTo Err_Handler
PushCallStack "mFunctions|GetRangeAddress"

If blnAbsoluteAddress Then
    GetRangeAddress = rng.Address
Else
    GetRangeAddress = rng.Address(False, False)
End If

Proc_Exit:
PopCallStack
Exit Function

Err_Handler:
GlobalErrHandler
Resume Proc_Exit
End Function

Public Function FirstDayInMonth(Optional dtmDate As Date = 0) As Date
    ' Return the first day in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    FirstDayInMonth = DateSerial(Year(dtmDate), _
     Month(dtmDate), 1)
End Function

Public Function LastDayInMonth(Optional dtmDate As Date = 0) As Date
    ' Return the last day in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    LastDayInMonth = DateSerial(Year(dtmDate), _
    Month(dtmDate) + 1, 0)
End Function

Public Function File_Dialog(dlgType As MsoFileDialogType, Optional strFilterDesc As String, Optional strFilters As String) As String
    Dim fd As FileDialog
    Dim i As Integer
    Dim strSelected As Variant

    Set fd = Application.FileDialog(dlgType)
       
    With fd
        If Len(strFilters) <> 0 Then
            .Filters.Clear
            .Filters.Add strFilterDesc, strFilters, 1
        End If
        
        .InitialFileName = ThisWorkbook.Path
        .AllowMultiSelect = False
    
        If .Show = -1 Then
            i = 0
            For Each strSelected In .SelectedItems
                File_Dialog = strSelected
            Next strSelected
        End If
        
    End With

    Set fd = Nothing
End Function
