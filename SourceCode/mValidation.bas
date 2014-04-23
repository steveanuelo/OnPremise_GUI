Attribute VB_Name = "mValidation"
Option Explicit

Public Function HasEmptyValues(ctrl As Control, Optional strName As String) As Boolean
    HasEmptyValues = False
    If Len(ctrl.Value) = 0 Or IsNull(ctrl.Value) Then
        If Len(strName) <> 0 Then MsgBox strName & " should not be empty.", vbExclamation, "Data Validation Error"
        HasEmptyValues = True
    End If
End Function

Public Function HasInvalidNumber(ctrl As Control, Optional IsOptional As Boolean = False, Optional strName As String) As Boolean
    HasInvalidNumber = False
    If IsOptional = False Or Len(ctrl.Value) <> 0 Then
        If IsNumeric(ctrl.Value) = False Then
            If Len(strName) <> 0 Then MsgBox strName & " should be numeric.", vbExclamation, "Data Validation Error"
            HasInvalidNumber = True
        End If
    End If
End Function

'Private Function Check_Number_Range(ctrl As Control, Optional strCondition As String, Optional strMsg As String) As Boolean
'
'If Not HasInvalidNumber(txtTT_GSV, True, "Trading Term % GSV") Then
'    If txtTT_GSV < 5 Then
'        MsgBox strMsg, vbOKOnly, "Data Validation Error"
'        txtTT_GSV.Text = vbNullString
'    End If
'End If
'
'End Function
