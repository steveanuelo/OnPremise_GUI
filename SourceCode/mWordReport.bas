Attribute VB_Name = "mWordReport"
Option Explicit

Public Const LONG_FORM_DOC_TEMP = "Long Form Template.docx"
Public Const SHORT_FORM_DOC_TEMP = "Short Form Template.docx"
Public Const TBL_QA3 = "QA3"
Public Const TBL_TERMS = "Terms"
Public Const TBL_COOP = "COOP"
Public Const TBL_AnP = "AnP"
Public Const TBL_COOP_AnP_Total = "COOP and AnP Total"
Public Const TBL_SUMMARY = "OP Summary"

Public Sub GenerateWordDocs(frm As frmMain)
      Dim objWord As Word.Application
      Dim doc As Word.Document
      Dim tbl As Word.Table
      Dim strDocTemplate As String
      Dim qry As String
      Dim rs As ADODB.Recordset
      Dim arr As Variant
      Dim i As Integer
      Dim arrTblToDel As Variant

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|GenerateWordDocs"

3     Set objWord = CreateObject("Word.Application")
4     objWord.Visible = True

5     Select Case frm.cboContractForm
          Case "Short Form"
6             strDocTemplate = SHORT_FORM_DOC_TEMP
7         Case "Long Form"
8             strDocTemplate = LONG_FORM_DOC_TEMP
9     End Select

10    objWord.Documents.Open Filename:=wkb.Path & "\" & strDocTemplate, ReadOnly:=True

11    Set doc = objWord.Documents(1)

12    Set rs = New ADODB.Recordset

13    ReDim arrTblToDel(0)

14    For i = 1 To doc.Tables.Count
15        Set tbl = doc.Tables(i)
          
16        qry = vbNullString
17        Select Case tbl.Title
              Case TBL_QA3
18                qry = "SELECT DISTINCT T1.ProductType, T2.BRAND_NAME, T2.PRODUCT_DESCRIPTION, " & _
                          "Format(T1.ContractedCases, '#,###'), " & _
                          "Format(T1.ContractedVolume, '#,###'), " & _
                          "Format(T1.ContractedGSV, '#,###'), " & _
                          "Format(T1.DirectPrice, '#,##0.00'), " & _
                          "Format(T1.WholesalePrice, '#,##0.00'), " & _
                          "Format(T1.QA3PerCaseUser + T1.QA3PerCaseAuto, '#,##0.00'), " & _
                          "Format(Round(T1.NIPOrLUCUser,2) + Round(T1.NIPOrLUCAuto, 2), '#,##0.00') " & _
                        "FROM " & OP_PROD_DETAILS_TBL & " AS T1 LEFT JOIN " & PRODUCT_MAP_TBL & " AS T2 " & _
                          "ON (T1.BrandCode = T2.BRAND_CODE) " & _
                         "AND (T1.ProductCode = T2.PRODUCT_CODE) " & _
                        "WHERE T1.RefNumber = '" & frm.txtRefNumber & "' " & _
                        "ORDER BY T1.ProductType, T2.BRAND_NAME " & _
                        "UNION ALL " & _
                        "SELECT 'TOTAL', '', '', " & _
                          "Format(SUM(T1.ContractedCases), '#,###'), " & _
                          "Format(SUM(T1.ContractedVolume), '#,###'), " & _
                          "Format(SUM(T1.ContractedGSV), '#,###'), '', '', '', '' " & _
                        "FROM " & OP_PROD_DETAILS_TBL & " AS T1 " & _
                        "WHERE T1.RefNumber = '" & frm.txtRefNumber & "'"
                        
19            Case TBL_TERMS
20                qry = "SELECT DISTINCT T1.ProductType, T2.BRAND_NAME, T2.PRODUCT_DESCRIPTION, " & _
                          "Format(T3.DollarPerLiter, '#,###'), " & _
                          "Format(T3.PctOfGSV, '#,###'), " & _
                          "Format(T3.FreqOfPayments, '#,###'), " & _
                          "Format(T3.AddnlDollarPerLiter, '#,###'), " & _
                          "Format(T3.AddnlPctOfGSV, '#,###'), " & _
                          "T3.CondTermComments " & _
                        "FROM (" & OP_PROD_DETAILS_TBL & " AS T1 LEFT JOIN " & PRODUCT_MAP_TBL & " AS T2 ON (T1.BrandCode = T2.BRAND_CODE) AND (T1.ProductCode = T2.PRODUCT_CODE)) " & _
                        " LEFT JOIN " & OP_TRADING_TERMS_TBL & " AS T3 ON (T1.RefNumber = T3.RefNumber) AND (T1.ProductCode = T3.ProductCode) " & _
                        "WHERE T1.RefNumber = '" & frm.txtRefNumber & "' " & _
                        "ORDER BY T1.ProductType, T2.BRAND_NAME "
21                If frm.chkNonContract.Value = -1 Then
22                    qry = qry & "UNION ALL " & _
                                  "SELECT '', '', 'NON CONTRACTED PRODUCTS', Format(T1.AllNonContrdProd_DollarperLtr, '#,###'), Format(T1.AllNonContrdProd_PctGSVlessQA3, '#,###'), '', '', '', '' " & _
                                  "FROM " & OP_TRADING_TERMS_CONST_TBL & " AS T1 " & _
                                  "WHERE T1.RefNumber = '" & frm.txtRefNumber & "'"
23                End If
24            Case TBL_COOP
25                qry = "SELECT T1.Amount " & _
                        "FROM ( " & _
                          "SELECT RefNumber, CashPaymentCoop AS Amount " & _
                          "FROM T_Main_COOP_And_AnP " & _
                          "UNION ALL " & _
                          "SELECT RefNumber, BonusStockCoop AS Amount " & _
                          "FROM T_Main_COOP_And_AnP " & _
                          "UNION ALL " & _
                          "SELECT RefNumber, PromoFundCoop AS Amount " & _
                          "FROM T_Main_COOP_And_AnP " & _
                          "UNION ALL " & _
                          "SELECT RefNumber, StaffIncentivesCoop AS Amount " & _
                          "FROM T_Main_COOP_And_AnP " & _
                          "UNION ALL " & _
                          "SELECT RefNumber, PRAHospitalityCoop AS Amount " & _
                          "FROM T_Main_COOP_And_AnP " & _
                        ") AS T1 " & _
                        "WHERE T1.RefNumber = '" & frm.txtRefNumber & "'"
26            Case TBL_AnP
27                qry = "SELECT T1.Amount " & _
                        "FROM ( " & _
                          "SELECT RefNumber, CashPaymentAnP AS Amount " & _
                          "FROM T_Main_COOP_And_AnP " & _
                          "UNION ALL " & _
                          "SELECT RefNumber, BonusStockAnP AS Amount " & _
                          "FROM T_Main_COOP_And_AnP " & _
                          "UNION ALL " & _
                          "SELECT RefNumber, PromoFundAnP AS Amount " & _
                          "FROM T_Main_COOP_And_AnP " & _
                          "UNION ALL " & _
                          "SELECT RefNumber, StaffIncentivesAnP AS Amount " & _
                          "FROM T_Main_COOP_And_AnP " & _
                          "UNION ALL " & _
                          "SELECT RefNumber, PRAHospitalityAnP AS Amount " & _
                          "FROM T_Main_COOP_And_AnP " & _
                        ") AS T1 " & _
                        "WHERE T1.RefNumber = '" & frm.txtRefNumber & "'"
28            Case TBL_COOP_AnP_Total
29                qry = "SELECT T1.Amount, T1.Comments " & _
                        "FROM ( " & _
                          "SELECT RefNumber, CashPaymentCoop + CashPaymentAnP AS Amount, CashPaymentComments AS Comments " & _
                          "FROM T_Main_COOP_And_AnP " & _
                          "UNION ALL " & _
                          "SELECT RefNumber, BonusStockCoop + BonusStockAnP AS Amount, BonusStockComments AS Comments " & _
                          "FROM T_Main_COOP_And_AnP " & _
                          "UNION ALL " & _
                          "SELECT RefNumber, PromoFundCoop + PromoFundAnP AS Amount, PromoFundComments AS Comments " & _
                          "FROM T_Main_COOP_And_AnP " & _
                          "UNION ALL " & _
                          "SELECT RefNumber, StaffIncentivesCoop + StaffIncentivesAnP AS Amount, StaffIncentivesComments AS Comments " & _
                          "FROM T_Main_COOP_And_AnP " & _
                          "UNION ALL " & _
                          "SELECT RefNumber, PRAHospitalityCoop + PRAHospitalityAnP AS Amount, PRAHospitalityComments AS Comments " & _
                          "FROM T_Main_COOP_And_AnP " & _
                          "UNION ALL " & _
                          "SELECT RefNumber, ReciprocalSpend AS Amount, ReciprocalSpendComments AS Comments " & _
                          "FROM T_Main_COOP_And_AnP " & _
                          "UNION ALL " & _
                          "SELECT RefNumber, CashPaymentCoop + CashPaymentAnP + BonusStockCoop + BonusStockAnP + PromoFundCoop + PromoFundAnP + StaffIncentivesCoop + StaffIncentivesAnP + PRAHospitalityCoop + PRAHospitalityAnP + ReciprocalSpend, '' " & _
                          "FROM T_Main_COOP_And_AnP " & _
                        ") AS T1 " & _
                        "WHERE T1.RefNumber = '" & frm.txtRefNumber & "'"
30            Case TBL_SUMMARY
31                qry = "SELECT DISTINCT T1.Family, T1.ProductType, T2.PRODUCT_DESCRIPTION, T1.ContractedCases, T1.ContractedGSV, T1.DirectPrice + T1.WholesalePrice, T1.QA3PerCaseUser + T1.QA3PerCaseAuto, Round(T1.NIPOrLUCUser,2) + Round(T1.NIPOrLUCAuto, 2) " & _
                        "FROM " & OP_PROD_DETAILS_TBL & " AS T1 LEFT JOIN " & PRODUCT_MAP_TBL & " AS T2 " & _
                          "ON (T1.BrandCode = T2.BRAND_CODE) " & _
                         "AND (T1.ProductCode = T2.PRODUCT_CODE) " & _
                        "WHERE T1.RefNumber = '" & frm.txtRefNumber & "' " & _
                        "ORDER BY T2.PRODUCT_DESCRIPTION " & _
                        "UNION ALL " & _
                        "SELECT '', '', 'TOTAL', SUM(T1.ContractedCases), SUM(T1.ContractedGSV), '', '', '' " & _
                        "FROM " & OP_PROD_DETAILS_TBL & " AS T1 " & _
                        "WHERE T1.RefNumber = '" & frm.txtRefNumber & "'"
32        End Select
          
33        If qry <> vbNullString Then
34            rs.Open qry, cn
              
35            If Not rs.EOF Then
36                arr = rs.GetRows(10000)
          
37                Select Case tbl.Title
                      Case TBL_QA3
38                        If IsTableEmpty(arr, 4) Then
39                            ReDim Preserve arrTblToDel(UBound(arrTblToDel) + 1)
40                            arrTblToDel(UBound(arrTblToDel)) = tbl.Title
41                        Else
42                            Call PopulateWordTable(tbl, arr)
43                        End If
44                    Case TBL_TERMS
45                        If IsTableEmpty(arr, 4) Then
46                            ReDim Preserve arrTblToDel(UBound(arrTblToDel) + 1)
47                            arrTblToDel(UBound(arrTblToDel)) = tbl.Title
48                        Else
49                            Call PopulateWordTable(tbl, arr, 3, 1)
50                        End If
      '                Case TBL_COOP
      '                    If IsTableEmpty(arr, 0) Then
      '                        ReDim Preserve arrTblToDel(UBound(arrTblToDel) + 1)
      '                        arrTblToDel(UBound(arrTblToDel)) = tbl.Title
      '                    Else
      '                        Call PopulateWordTable(tbl, arr, 2, 2, False, False, True)
      '                    End If
      '                Case TBL_AnP
      '                    If IsTableEmpty(arr, 0) Then
      '                        ReDim Preserve arrTblToDel(UBound(arrTblToDel) + 1)
      '                        arrTblToDel(UBound(arrTblToDel)) = tbl.Title
      '                    Else
      '                        Call PopulateWordTable(tbl, arr, 2, 2, False, False, True)
      '                    End If
51                    Case TBL_COOP_AnP_Total
52                        If IsTableEmpty(arr, 0) Then
53                            ReDim Preserve arrTblToDel(UBound(arrTblToDel) + 1)
54                            arrTblToDel(UBound(arrTblToDel)) = tbl.Title
55                        Else
56                            Call PopulateWordTable(tbl, arr, 2, 2, False, False, True)
57                        End If
      '                Case TBL_SUMMARY
      '                    If IsTableEmpty(arr, 3) Then
      '                        ReDim Preserve arrTblToDel(UBound(arrTblToDel) + 1)
      '                        arrTblToDel(UBound(arrTblToDel)) = tbl.Title
      '                    Else
      '                        Call PopulateWordTable(tbl, arr, 3, 1)
      '                    End If
58                End Select
          
59            End If
60            Call CloseRecordset(rs)
61        End If
62    Next i

      ' Delete empty tables
63    For Each tbl In doc.Tables
64        For i = 1 To UBound(arrTblToDel)
65            If arrTblToDel(i) = tbl.Title Then
66                tbl.Delete
67                Exit For
68            End If
69        Next i
70    Next tbl

71    MsgBox "Finished generating contracts.", vbInformation
72    objWord.Activate

      'doc.Close False
73    Set doc = Nothing
      'objWord.Quit
74    Set objWord = Nothing

Proc_Exit:
75    PopCallStack
76    Exit Sub

Err_Handler:
77    GlobalErrHandler
78    Resume Proc_Exit
End Sub

Private Function IsTableEmpty(arr As Variant, intStartCol As Integer) As Boolean
      Dim x As Integer, y As Integer
      Dim dblSum As Double

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mWordReport|IsTableEmpty"

3     IsTableEmpty = False

4     dblSum = 0
5     For x = 0 To UBound(arr, 2)
6         For y = intStartCol To UBound(arr, 1)
7             If IsNumeric(arr(y, x)) Then
8                 dblSum = dblSum + CDbl(arr(y, x))
9             End If
10        Next y
11    Next x

12    If dblSum = 0 Then
13        IsTableEmpty = True
14    End If

Proc_Exit:
15    PopCallStack
16    Exit Function

Err_Handler:
17    GlobalErrHandler
18    Resume Proc_Exit

End Function

Private Sub PopulateWordTable(ByRef tbl As Word.Table, arr As Variant, _
                              Optional intStartRow As Integer = 2, Optional intStartCol As Integer = 1, _
                              Optional blnAddRows As Boolean = True, Optional blnDeleteBlankCols As Boolean = True, Optional blnDeleteBlankRows As Boolean = False)
      Dim x As Integer, y As Integer
      Dim blnEmptyCol As Boolean
      Dim blnEmptyRow As Boolean
      Dim varText As Variant

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mWordReport|PopulateWordTable"

3     For x = 0 To UBound(arr, 2)
4         For y = 0 To UBound(arr, 1)
5             tbl.cell(x + intStartRow, y + intStartCol).Range.Text = SetEmptyValue(Trim(arr(y, x)), NullStrings)
6         Next y
          
          ' Insert new rows
7         If blnAddRows Then
8             If x <> UBound(arr, 2) Then tbl.Rows.Add
9         End If
10    Next x

      ' Delete empty Columns
11    If blnDeleteBlankCols Then
12        For y = UBound(arr, 1) To 0 Step -1
              
13            blnEmptyCol = True
              
14            For x = 0 To UBound(arr, 2)
15                varText = SetEmptyValue(Mid(tbl.cell(x + intStartRow, y + intStartCol).Range.Text, 1, Len(tbl.cell(x + intStartRow, y + intStartCol).Range.Text) - 2))
16                If IsNumeric(varText) Then
17                    varText = CLng(varText)
18                End If
                  
19                If Not (CStr(varText) = "0" Or _
                          Len(varText) = 0) Then
20                    blnEmptyCol = False
21                    Exit For
22                End If
23            Next x
              
24            If blnEmptyCol Then tbl.Columns(y + 1).Delete
25        Next y
26    End If

      ' Delete empty rows
27    If blnDeleteBlankRows Then
28        For x = UBound(arr, 2) To 0 Step -1
29            blnEmptyRow = True
              
30            For y = 0 To UBound(arr, 1)
31                varText = SetEmptyValue(Mid(tbl.cell(x + intStartRow, y + intStartCol).Range.Text, 1, Len(tbl.cell(x + intStartRow, y + intStartCol).Range.Text) - 2))
32                If IsNumeric(varText) Then
33                    varText = CLng(varText)
34                End If
                  
35                If Not (CStr(varText) = "0" Or _
                          Len(varText) = 0) Then
36                    blnEmptyRow = False
37                    Exit For
38                End If
39            Next y
              
40            If blnEmptyRow Then tbl.Rows(x + intStartRow).Delete
41        Next x
42    End If


Proc_Exit:
43    PopCallStack
44    Exit Sub

Err_Handler:
45    GlobalErrHandler
46    Resume Proc_Exit
End Sub
