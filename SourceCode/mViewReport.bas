Attribute VB_Name = "mViewReport"
Option Explicit

Dim intCM_Row As Integer

Public Function CreateViewWorkbook(arrSheet As Variant) As Workbook
      Dim xlApp As Excel.Application
      Dim wb As Workbook
      Dim ws As Worksheet
      Dim strFile As String
      Dim i As Integer

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mViewReport|CreateViewWorkbook"

3     Application.DisplayAlerts = False

      ' Open a new workbook
4     Set wb = Application.Workbooks.Add

      ' Copy template sheets
5     For i = 0 To UBound(arrSheet)
6         ThisWorkbook.Worksheets(arrSheet(i)).Copy Before:=wb.Worksheets("Sheet1")
7     Next i

      ' Delete extra sheets
8     For Each ws In wb.Worksheets
9         If Left(ws.Name, 5) = "Sheet" Then ws.Delete
10    Next ws

11    wb.Activate

12    Set CreateViewWorkbook = wb

13    Application.DisplayAlerts = True

14    Set wb = Nothing

Proc_Exit:
15    PopCallStack
16    Exit Function

Err_Handler:
17    GlobalErrHandler
18    Resume Proc_Exit
End Function

Public Sub DeleteViewTemplates(wb As Workbook, arrSheet As Variant)
      Dim i As Integer

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mViewReport|DeleteViewTemplates"

3     wb.Application.DisplayAlerts = False

4     For i = 0 To UBound(arrSheet)
5         wb.Worksheets(arrSheet(i)).Delete
6     Next i

7     wb.Application.Visible = True

8     wb.Application.DisplayAlerts = True

Proc_Exit:
9     PopCallStack
10    Exit Sub

Err_Handler:
11    GlobalErrHandler
12    Resume Proc_Exit
End Sub

Public Sub PEM_Preview(wb As Workbook, frm As frmMain)
      Dim i As Integer, j As Integer
      Dim arrData As Variant
      Dim arrProducts As Variant
      Dim arrTerms As Variant
      Dim arrQA3 As Variant
      Dim dblTotalGSV As Double

      Const COL_ProdDesc = 0
      Const COL_ProdCode = 1
      Const COL_ContrdVol = 2
      Const COL_ContrdGSV = 3
      Const COL_BannerTerms = 4
      Const COL_StandardTerms = 5
      Const COL_AddnlTerms = 6
      Const COL_KWI = 7
      Const COL_COP = 8
      Const COL_QA3 = 9
      Const COL_COOP = 10
      Const COL_AnD = 11
      Const COL_NSV = 12
      Const COL_COGSnDistr = 13
      Const COL_CM = 14
      Const COL_AnDPerGSV = 15
      Const COL_NSVPerLtr = 16
      Const COL_CMPerNSV = 17
      Const COL_LUC = 18
      Const COL_NIP = 19


      Const DECIMAL_0_FORMAT = "#,##0"
      Const DECIMAL_2_FORMAT = "#,##0.00"
      Const PERCENT_FORMAT = "0.0"

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mViewReport|PEM_Preview"

      '' Clear listbox data
      'frm.lstPEMPreview.Clear

3     arrProducts = frm.lstProducts.List
4     arrTerms = frm.lstTrdTerms.List
5     arrQA3 = frm.lstQA3.List

      ' Populate array
6     ReDim arrData(frm.lstProducts.ListCount - 1, frm.lstPEMPreview.ColumnCount - 1)

7     If frm.lstProducts.ListCount <> 0 Then
          ' Get total GSV
8         dblTotalGSV = 0
9         For i = 0 To frm.lstProducts.ListCount - 1
10            dblTotalGSV = dblTotalGSV + frm.lstProducts.List(i, ProdList_ContractGSV)
11        Next i

12        For i = 0 To frm.lstProducts.ListCount - 1
13            arrData(i, COL_ProdDesc) = arrProducts(i, ProdList_ProdDesc)
14            arrData(i, COL_ProdCode) = arrProducts(i, ProdList_ProdCode)
15            arrData(i, COL_ContrdVol) = ConvToDbl(arrProducts(i, ProdList_ContractVol))
16            arrData(i, COL_ContrdGSV) = ConvToDbl(arrProducts(i, ProdList_ContractGSV))
17            arrData(i, COL_QA3) = ConvToDbl(arrQA3(i, QA3List_QA3))
18            arrData(i, COL_BannerTerms) = ConvToDbl(arrTerms(i, TTList_BannerTerm))
19            arrData(i, COL_StandardTerms) = ConvToDbl(arrTerms(i, TTList_StandardTerm))
20            arrData(i, COL_AddnlTerms) = ConvToDbl(arrTerms(i, TTList_AddnlTerm))
21            arrData(i, COL_KWI) = ConvToDbl(arrQA3(i, QA3List_KWI))
22            arrData(i, COL_COP) = ConvToDbl(arrQA3(i, QA3List_COP))
23            arrData(i, COL_COOP) = (ConvToDbl(frm.txtCoopCashPay) + ConvToDbl(frm.txtCoopBonusStock) + ConvToDbl(frm.txtCoopPromoFund) + ConvToDbl(frm.txtCoopStaffIncentives) + ConvToDbl(frm.txtCoopPRAHospitality)) * (arrData(i, COL_ContrdGSV) / dblTotalGSV)
24            arrData(i, COL_AnD) = ConvToDbl(arrData(i, COL_QA3)) + ConvToDbl(arrData(i, COL_BannerTerms)) + ConvToDbl(arrData(i, COL_StandardTerms)) + ConvToDbl(arrData(i, COL_AddnlTerms)) + ConvToDbl(arrData(i, COL_KWI)) + ConvToDbl(arrData(i, COL_COP)) + ConvToDbl(arrData(i, COL_COOP))
25            arrData(i, COL_NSV) = arrData(i, COL_ContrdGSV) - arrData(i, COL_AnD)
26            arrData(i, COL_COGSnDistr) = ConvToDbl(arrData(i, COL_ContrdVol)) * GetItemFromMappingTbl(COGSPERLTR_MAP_TBL, "COGSperLitre", strWhereCondit:="ProductCode = """ & CStr(arrData(i, COL_ProdCode)) & """ AND Start_Date <=#" & GetPromoDate(End_Date, frm) & "# AND End_Date >=#" & GetPromoDate(Start_Date, frm) & "#")
27            arrData(i, COL_CM) = arrData(i, COL_NSV) - arrData(i, COL_COGSnDistr)
28            arrData(i, COL_AnDPerGSV) = (arrData(i, COL_AnD) / arrData(i, COL_ContrdGSV)) * 100
29            arrData(i, COL_NSVPerLtr) = arrData(i, COL_NSV) / arrData(i, COL_ContrdVol)
30            arrData(i, COL_CMPerNSV) = (arrData(i, COL_CM) / arrData(i, COL_NSV)) * 100
31            arrData(i, COL_NIP) = 0
32            arrData(i, COL_LUC) = 0
33            Select Case arrQA3(i, QA3List_Family)
                  Case "SPIRITS"
34                    arrData(i, COL_NIP) = ConvToDbl(arrQA3(i, QA3List_NipOrLUCInput)) + ConvToDbl(arrQA3(i, QA3List_NipOrLUCAuto))
35                Case "WINE"
36                    arrData(i, COL_LUC) = ConvToDbl(arrQA3(i, QA3List_NipOrLUCInput)) + ConvToDbl(arrQA3(i, QA3List_NipOrLUCAuto))
37                Case "OTHER ALCOHOLIC BEVERAGES"
38                    arrData(i, COL_LUC) = ConvToDbl(arrQA3(i, QA3List_NipOrLUCInput)) + ConvToDbl(arrQA3(i, QA3List_NipOrLUCAuto))
39            End Select
40        Next i
41    End If

      ' Calculate Totals
42    frm.txtPEM_Total_Vol = Format(SumArrayFields(arrData, COL_ContrdVol), DECIMAL_0_FORMAT)
43    frm.txtPEM_Total_GSV = Format(SumArrayFields(arrData, COL_ContrdGSV), DECIMAL_0_FORMAT)
44    frm.txtPEM_Total_BannerTerms = Format(SumArrayFields(arrData, COL_BannerTerms), DECIMAL_0_FORMAT)
45    frm.txtPEM_Total_StanTerms = Format(SumArrayFields(arrData, COL_StandardTerms), DECIMAL_0_FORMAT)
46    frm.txtPEM_Total_AddTerms = Format(SumArrayFields(arrData, COL_AddnlTerms), DECIMAL_0_FORMAT)
47    frm.txtPEM_Total_KWI = Format(SumArrayFields(arrData, COL_KWI), DECIMAL_0_FORMAT)
48    frm.txtPEM_Total_COP = Format(SumArrayFields(arrData, COL_COP), DECIMAL_0_FORMAT)
49    frm.txtPEM_Total_QA3 = Format(SumArrayFields(arrData, COL_QA3), DECIMAL_0_FORMAT)
50    frm.txtPEM_Total_COOP = Format(SumArrayFields(arrData, COL_COOP), DECIMAL_0_FORMAT)
51    frm.txtPEM_Total_AnD = Format(SumArrayFields(arrData, COL_AnD), DECIMAL_0_FORMAT)
52    frm.txtPEM_Total_NSV = Format(SumArrayFields(arrData, COL_NSV), DECIMAL_0_FORMAT)
53    frm.txtPEM_Total_COGS = Format(SumArrayFields(arrData, COL_COGSnDistr), DECIMAL_0_FORMAT)
54    frm.txtPEM_Total_CM = Format(SumArrayFields(arrData, COL_CM), DECIMAL_0_FORMAT)
55    frm.txtPEM_Total_AnD_GSV = Format((ConvToDbl(frm.txtPEM_Total_AnD) / ConvToDbl(frm.txtPEM_Total_GSV) * 100), PERCENT_FORMAT)
56    frm.txtPEM_Total_NSV_Per_Ltr = Format((ConvToDbl(frm.txtPEM_Total_NSV) / ConvToDbl(frm.txtPEM_Total_Vol)), DECIMAL_2_FORMAT)
57    frm.txtPEM_Total_CM_NSV = Format((ConvToDbl(frm.txtPEM_Total_CM) / ConvToDbl(frm.txtPEM_Total_NSV) * 100), PERCENT_FORMAT)
58    frm.txtPEM_Total_LUC = Format(SumArrayFields(arrData, COL_LUC), DECIMAL_2_FORMAT)
59    frm.txtPEM_Total_NIP = Format(SumArrayFields(arrData, COL_NIP), DECIMAL_2_FORMAT)

      ' Format data
60    For i = 0 To UBound(arrData)
61        For j = 2 To UBound(arrData, 2)
62            If j = COL_AnDPerGSV Or j = COL_CMPerNSV Then
63                arrData(i, j) = Format(arrData(i, j), PERCENT_FORMAT)
64            ElseIf j = COL_NSVPerLtr Or j = COL_NIP Or j = COL_LUC Then
65                arrData(i, j) = Format(arrData(i, j), DECIMAL_2_FORMAT)
66            Else
67                arrData(i, j) = Format(arrData(i, j), DECIMAL_0_FORMAT)
68            End If
69        Next j
70    Next i

      ' Populate listbox
71    If frm.lstProducts.ListCount > 0 Then
72        frm.lstPEMPreview.List = arrData
73    End If

Proc_Exit:
74    PopCallStack
75    Exit Sub

Err_Handler:
76    GlobalErrHandler
77    Resume Proc_Exit
End Sub

Private Function SumArrayFields(arr As Variant, intField As Integer) As Double
      Dim i As Integer
      Dim dblSum As Double

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mViewReport|SumArrayFields"

3     SumArrayFields = 0

4     For i = 0 To UBound(arr)
5         dblSum = dblSum + arr(i, intField)
6     Next i

7     SumArrayFields = dblSum

Proc_Exit:
8     PopCallStack
9     Exit Function

Err_Handler:
10    GlobalErrHandler
11    Resume Proc_Exit

End Function

Public Sub CreatePEMReport(wb As Workbook, frm As frmMain)
      Dim ws As Worksheet
      Dim rng As Range
      Dim intRowCount As Integer
      Dim rs As ADODB.Recordset
      Dim qry As String
      Dim strRefNumber As String
      Dim i As Integer

      Const COP_PCT = 0.5

      Const RNG_OP_REF = "C1"
      Const RNG_CUST_NAME = "C2"
      Const RNG_CONTRACT_START = "C3"
      Const RNG_CONTRACT_END = "C4"
      Const RNG_MANAGER = "C5"

      Const RNG_START = "D8"
      Const ROW_FAMILY = -2
      Const ROW_PRODUCTTYPE = -1
      Const ROW_PRODUCT = 0
      Const ROW_VOL = 3
      Const ROW_VOL9L = 4
      Const ROW_GSV = 6
      Const ROW_BANNERTERMS = 8
      Const ROW_STANDTERMS = 9
      Const ROW_ADDNLTERMS = 10
      Const ROW_KWI = 11
      Const ROW_COP = 12
      Const ROW_QA3 = 13
      Const ROW_COOP = 14
      Const ROW_ALLOWnDISC = 15
      Const ROW_NETSALES = 17
      Const ROW_COGSnDIST = 19
      Const ROW_CONTRIBMARG = 21
      Const ROW_AnP = 23
      Const ROW_CAAP = 25
      Const ROW_GSVperVOL = 28
      Const ROW_ALLOWnDISCperGSV = 29
      Const ROW_NETSALESperVOL = 30
      Const ROW_COGSnDISTperVOL = 31
      Const ROW_CONTRIBMARGperVOL = 32
      Const ROW_CONTRIBMARGperNETSALES = 33
      Const ROW_AnPperNETSALES = 34
      Const ROW_CAAPperVOL = 35
      Const ROW_TOTTERMS = 37
      Const ROW_QA3per9lc = 38
      Const ROW_LUCperNIP = 39
      Const ROW_LUC = 40
      Const ROW_NIP = 41

      Const ROW_CashPaymentCoop = 43
      Const ROW_BonusStockCoop = 44
      Const ROW_PromoFundCoop = 45
      Const ROW_StaffIncentivesCoop = 46
      Const ROW_PRAHospitalityCoop = 47
      Const ROW_CashPaymentAnP = 50
      Const ROW_BonusStockAnP = 51
      Const ROW_PromoFundAnP = 52
      Const ROW_StaffIncentivesAnP = 53
      Const ROW_PRAHospitalityAnP = 54
      Const ROW_ReciprocalSpend = 57

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mViewReport|CreatePurchaseOrder"

3     wb.Application.DisplayAlerts = False
        
      ' Copy sheet template
4     wb.Worksheets(PEM_TEMP_SHEET).Copy After:=wb.Worksheets(wb.Worksheets.Count)
5     Set ws = wb.ActiveSheet

      ' Rename sheet
6     ws.Name = PEM_TEMP_SHEET_RENAME

      ' Get Ref number
7     strRefNumber = frm.txtRefNumber

      ' Populate table
8     Set rs = New ADODB.Recordset

9     qry = "SELECT DISTINCT T1.RefNumber, T2.ProductCode, T2.Family, T2.ProductType, T3.PRODUCT_DESCRIPTION, T1.RouteToMarket, T1.Wholesaler,T2.ContractedCases, " & _
              "T2.UnitsPerCase, T2.ContractedVolume AS [Volume], T2.ContractedGSV AS [GSV], T5.StandardTerms, T5.AdditionalTerms, T5.BannerTerms, T2.KWI, T2.COP, T5.DollarPerLiter AS [Terms_P/L], T5.PctOfGSV AS [Terms_%GSV], " & _
              "T2.QA3, IIF(T2.NIPOrLUCUser=0,T2.NIPOrLUCAuto,T2.NIPOrLUCUser) AS NIPPrice, " & _
              "(T6.CashPaymentCoop + T6.BonusStockCoop + T6.PromoFundCoop + T6.StaffIncentivesCoop + T6.PRAHospitalityCoop) * (T2.ContractedGSV / T7.Total_GSV) AS COOP, " & _
              "(T6.CashPaymentAnP + T6.BonusStockAnP + T6.PromoFundAnP + T6.StaffIncentivesAnP + T6.PRAHospitalityAnP) * (T2.ContractedGSV / T7.Total_GSV) AS [A&P], " & _
              "IIF(T2.Family='WINE' OR T2.Family='OTHER ALCOHOLIC BEVERAGES',IIF(T2.NIPOrLUCAuto=0,T2.NIPOrLUCUser,T2.NIPOrLUCAuto),0) AS [LUC], " & _
              "IIF(T2.Family='SPIRITS',IIF(T2.NIPOrLUCAuto=0,T2.NIPOrLUCUser,T2.NIPOrLUCAuto),0) AS [NIP], " & _
              "T6.ReciprocalSpend, (T4.COGSperLitre * T2.ContractedVolume) AS [COGSandDist] " & _
            "FROM (((((" & OP_MAIN_TBL & " AS T1 " & _
              "INNER JOIN " & OP_PROD_DETAILS_TBL & " AS T2 ON T1.RefNumber = T2.RefNumber) " & _
              "INNER JOIN " & PRODUCT_MAP_TBL & " AS T3 ON T2.ProductCode = T3.PRODUCT_CODE) " & _
              "INNER JOIN " & COGSPERLTR_MAP_TBL & " AS T4 ON T3.PRODUCT_CODE = T4.ProductCode) " & _
              "INNER JOIN " & OP_TRADING_TERMS_TBL & " AS T5 ON T2.RefNumber = T5.RefNumber AND T2.ProductCode = T5.ProductCode) " & _
              "INNER JOIN " & OP_COOP_ANP_TBL & " AS T6 ON T2.RefNumber = T6.RefNumber) " & _
              "INNER JOIN (SELECT RefNumber, SUM(ContractedGSV) AS Total_GSV FROM " & OP_PROD_DETAILS_TBL & " GROUP BY RefNumber) AS T7 ON T2.RefNumber = T7.RefNumber " & _
            "WHERE T1.RefNumber = '" & strRefNumber & "' " & _
              "AND T4.Start_Date <=#" & GetPromoDate(End_Date, frm) & "# AND T4.End_Date >=#" & GetPromoDate(Start_Date, frm) & "# " & _
            "ORDER BY T3.PRODUCT_DESCRIPTION"
10    rs.Open qry, cn

      ' Generate report
11    If Not rs.EOF Then
          ' Print headers
12        ws.Range(RNG_OP_REF).Value = frm.txtRefNumber.Value
13        ws.Range(RNG_CUST_NAME).Value = frm.txtOutletOrGroupName
14        ws.Range(RNG_CONTRACT_START).Value = frm.txtFromDate.Text
15        ws.Range(RNG_CONTRACT_END).Value = frm.txtToDate.Text
16        ws.Range(RNG_MANAGER).Value = frm.cboCreator.Text

          ' Create columns
17        rs.MoveFirst
18        intRowCount = UBound(rs.GetRows(5000), 2)
19        For i = 0 To intRowCount - 1
20            ws.Range(RNG_START).EntireColumn.Insert xlShiftToRight
21        Next i
22        rs.MoveFirst
          
          ' Set starting range
23        Set rng = ws.Range(RNG_START)
          
          ' Print data
24        i = 0
25        While Not rs.EOF
26            With rng
                  ' Family (hidden)
27                .Offset(ROW_FAMILY, i).Value = rs.Fields("Family").Value
                  ' Product Type (hidden)
28                .Offset(ROW_PRODUCTTYPE, i).Value = rs.Fields("ProductType").Value
                  ' Product Description
29                .Offset(ROW_PRODUCT, i).Value = rs.Fields("PRODUCT_DESCRIPTION").Value
                  ' Volume
30                .Offset(ROW_VOL, i).Value = rs.Fields("Volume").Value
                  ' Volume 9L cases = Volume / 9
31                .Offset(ROW_VOL9L, i).Formula = "=" & GetRangeAddress(.Offset(ROW_VOL, i)) & "/9"
                  ' Gross Sales
32                .Offset(ROW_GSV, i).Value = rs.Fields("GSV").Value
                  ' Banner Terms (negate)
33                .Offset(ROW_BANNERTERMS, i).Value = -1 * rs.Fields("BannerTerms").Value
                  ' Standards Terms (negate)
34                .Offset(ROW_STANDTERMS, i).Value = -1 * rs.Fields("StandardTerms").Value
                  ' Additional Terms (negate)
35                .Offset(ROW_ADDNLTERMS, i).Value = -1 * rs.Fields("AdditionalTerms").Value
                  ' KWI (negate)
36                .Offset(ROW_KWI, i).Value = -1 * rs.Fields("KWI").Value
                  ' COP (negate)
37                .Offset(ROW_COP, i).Value = -1 * rs.Fields("COP").Value
                  ' QA3 (negate)
38                .Offset(ROW_QA3, i).Value = -1 * rs.Fields("QA3").Value
                  ' COOP (negate)
39                .Offset(ROW_COOP, i).Value = -1 * rs.Fields("COOP").Value
                  ' Allowance and Discount = Terms + KWI + COP + QA3 + COOP
40                .Offset(ROW_ALLOWnDISC, i).Formula = "=SUM(" & GetRangeAddress(.Offset(ROW_BANNERTERMS, i)) & ":" & GetRangeAddress(.Offset(ROW_COOP, i)) & ")"
                  ' Net Sales = GSV + Allowance and Discount
41                .Offset(ROW_NETSALES, i).Formula = "=" & GetRangeAddress(.Offset(ROW_GSV, i)) & "+" & GetRangeAddress(.Offset(ROW_ALLOWnDISC, i))
                  ' COGS and Dist
42                .Offset(ROW_COGSnDIST, i).Value = -1 * rs.Fields("COGSandDist").Value
                  ' Contributive Margin = Net Sales + COGS and Dist
43                .Offset(ROW_CONTRIBMARG, i).Formula = "=" & GetRangeAddress(.Offset(ROW_NETSALES, i)) & "+" & GetRangeAddress(.Offset(ROW_COGSnDIST, i))
                  ' A&P
44                .Offset(ROW_AnP, i).Value = -1 * rs.Fields("A&P").Value
                  ' CAAP = Contributive Margin + A&P
45                .Offset(ROW_CAAP, i).Formula = "=" & GetRangeAddress(.Offset(ROW_CONTRIBMARG, i)) & "+" & GetRangeAddress(.Offset(ROW_AnP, i))
                  
                  ' Gross Sales/L = Gross Sales / Volume
46                .Offset(ROW_GSVperVOL, i).Formula = "=IFERROR(" & GetRangeAddress(.Offset(ROW_GSV, i)) & "/" & GetRangeAddress(.Offset(ROW_VOL, i)) & ", 0)"
                  ' Allowance&Discount % Gross Sales = Allowance and Discount / Gross Sales)
47                .Offset(ROW_ALLOWnDISCperGSV, i).Formula = "=IFERROR(-" & GetRangeAddress(.Offset(ROW_ALLOWnDISC, i)) & "/" & GetRangeAddress(.Offset(ROW_GSV, i)) & ", 0)"
                  ' Net Sales/L = Net Sales / Volume
48                .Offset(ROW_NETSALESperVOL, i).Formula = "=IFERROR(" & GetRangeAddress(.Offset(ROW_NETSALES, i)) & "/" & GetRangeAddress(.Offset(ROW_VOL, i)) & ", 0)"
                  ' COGS&Dist/L = (COGS and Dist / Volume)
49                .Offset(ROW_COGSnDISTperVOL, i).Formula = "=IFERROR(-" & GetRangeAddress(.Offset(ROW_COGSnDIST, i)) & "/" & GetRangeAddress(.Offset(ROW_VOL, i)) & ", 0)"
                  ' Contributive Margin/L = Contributive Margin / Volume
50                .Offset(ROW_CONTRIBMARGperVOL, i).Formula = "=IFERROR(" & GetRangeAddress(.Offset(ROW_CONTRIBMARG, i)) & "/" & GetRangeAddress(.Offset(ROW_VOL, i)) & ", 0)"
                  ' Contributive Margin/NSV = Contributive Margin / Net Sales
51                .Offset(ROW_CONTRIBMARGperNETSALES, i).Formula = "=IFERROR(" & GetRangeAddress(.Offset(ROW_CONTRIBMARG, i)) & "/" & GetRangeAddress(.Offset(ROW_NETSALES, i)) & ", 0)"
                  ' Total A&P % Net Sales = (A&P / Net Sales)
52                .Offset(ROW_AnPperNETSALES, i).Formula = "=IFERROR(-" & GetRangeAddress(.Offset(ROW_AnP, i)) & "/" & GetRangeAddress(.Offset(ROW_NETSALES, i)) & ", 0)"
                  ' CAAP/L = CAAP / Volume
53                .Offset(ROW_CAAPperVOL, i).Formula = "=IFERROR(" & GetRangeAddress(.Offset(ROW_CAAP, i)) & "/" & GetRangeAddress(.Offset(ROW_VOL, i)) & ", 0)"
                  
                  ' Total Terms % GSV = -(Standard Terms + Additional Terms) / GSV
54                .Offset(ROW_TOTTERMS, i).Formula = "=IFERROR(-(" & GetRangeAddress(.Offset(ROW_STANDTERMS, i)) & "+" & GetRangeAddress(.Offset(ROW_ADDNLTERMS, i)) & "" & ")/" & GetRangeAddress(.Offset(ROW_GSV, i)) & ", 0)"
                  ' QA3 per 9lc = -QA3 / Volume 9L
55                .Offset(ROW_QA3per9lc, i).Formula = "=IFERROR(-" & GetRangeAddress(.Offset(ROW_QA3, i)) & "/" & GetRangeAddress(.Offset(ROW_VOL9L, i)) & ", 0)"
                  ' LUC
56                .Offset(ROW_LUC, i).Formula = rs.Fields("LUC").Value
                  ' NIP
57                .Offset(ROW_NIP, i).Formula = rs.Fields("NIP").Value
                  ' LUC / NIP
58                .Offset(ROW_LUCperNIP, i).Formula = rs.Fields("LUC").Value + rs.Fields("NIP").Value
                  
59            End With 'rng
              
60            rs.MoveNext
61            i = i + 1
62        Wend
          
          ' Set Totals formula
63        With rng
              ' Volume
64            .Offset(ROW_VOL, -1).Formula = "=SUM(" & GetRangeAddress(.Offset(ROW_VOL, 0)) & ":" & GetRangeAddress(.Offset(ROW_VOL, intRowCount)) & ")"
              ' Total Gross Sales
65            .Offset(ROW_GSV, -1).Formula = "=SUM(" & GetRangeAddress(.Offset(ROW_GSV, 0)) & ":" & GetRangeAddress(.Offset(ROW_GSV, intRowCount)) & ")"
              ' Banner Terms
66            .Offset(ROW_BANNERTERMS, -1).Formula = "=SUM(" & GetRangeAddress(.Offset(ROW_BANNERTERMS, 0)) & ":" & GetRangeAddress(.Offset(ROW_BANNERTERMS, intRowCount)) & ")"
              ' Standard Terms
67            .Offset(ROW_STANDTERMS, -1).Formula = "=SUM(" & GetRangeAddress(.Offset(ROW_STANDTERMS, 0)) & ":" & GetRangeAddress(.Offset(ROW_STANDTERMS, intRowCount)) & ")"
              ' Additional Terms
68            .Offset(ROW_ADDNLTERMS, -1).Formula = "=SUM(" & GetRangeAddress(.Offset(ROW_ADDNLTERMS, 0)) & ":" & GetRangeAddress(.Offset(ROW_ADDNLTERMS, intRowCount)) & ")"
              ' KWI
69            .Offset(ROW_KWI, -1).Formula = "=SUM(" & GetRangeAddress(.Offset(ROW_KWI, 0)) & ":" & GetRangeAddress(.Offset(ROW_KWI, intRowCount)) & ")"
              ' COP
70            .Offset(ROW_COP, -1).Formula = "=SUM(" & GetRangeAddress(.Offset(ROW_COP, 0)) & ":" & GetRangeAddress(.Offset(ROW_COP, intRowCount)) & ")"
              ' QA3
71            .Offset(ROW_QA3, -1).Formula = "=SUM(" & GetRangeAddress(.Offset(ROW_QA3, 0)) & ":" & GetRangeAddress(.Offset(ROW_QA3, intRowCount)) & ")"
              ' COOP
72            .Offset(ROW_COOP, -1).Formula = "=SUM(" & GetRangeAddress(.Offset(ROW_COOP, 0)) & ":" & GetRangeAddress(.Offset(ROW_COOP, intRowCount)) & ")"
              ' COGS and Dist
73            .Offset(ROW_COGSnDIST, -1).Formula = "=SUM(" & GetRangeAddress(.Offset(ROW_COGSnDIST, 0)) & ":" & GetRangeAddress(.Offset(ROW_COGSnDIST, intRowCount)) & ")"
              ' A&P
74            .Offset(ROW_AnP, -1).Formula = "=SUM(" & GetRangeAddress(.Offset(ROW_AnP, 0)) & ":" & GetRangeAddress(.Offset(ROW_AnP, intRowCount)) & ")"
              ' LUC
75            .Offset(ROW_LUC, -1).Formula = "=SUM(" & GetRangeAddress(.Offset(ROW_LUC, 0)) & ":" & GetRangeAddress(.Offset(ROW_LUC, intRowCount)) & ")"
              ' NIP
76            .Offset(ROW_NIP, -1).Formula = "=SUM(" & GetRangeAddress(.Offset(ROW_NIP, 0)) & ":" & GetRangeAddress(.Offset(ROW_NIP, intRowCount)) & ")"
77        End With 'rng

78    End If
79    Call CloseRecordset(rs)

      ' Calculate formulas
80    Application.Calculate


      ' COOP and A&P split totals
81    qry = "SELECT CashPaymentCoop, BonusStockCoop, PromoFundCoop, StaffIncentivesCoop, PRAHospitalityCoop, " & _
              "CashPaymentAnP, BonusStockAnP, PromoFundAnP, StaffIncentivesAnP, PRAHospitalityAnP, " & _
              "ReciprocalSpend " & _
            "FROM " & OP_COOP_ANP_TBL & " " & _
            "WHERE RefNumber = '" & strRefNumber & "'"
82    rs.Open qry, cn

83    If Not rs.EOF Then
84        With rng
              ' COOP
85            .Offset(ROW_CashPaymentCoop, -2).Value = rs.Fields("CashPaymentCoop").Value
86            .Offset(ROW_BonusStockCoop, -2).Value = rs.Fields("BonusStockCoop").Value
87            .Offset(ROW_PromoFundCoop, -2).Value = rs.Fields("PromoFundCoop").Value
88            .Offset(ROW_StaffIncentivesCoop, -2).Value = rs.Fields("StaffIncentivesCoop").Value
89            .Offset(ROW_PRAHospitalityCoop, -2).Value = rs.Fields("PRAHospitalityCoop").Value
          
              ' A&P
90            .Offset(ROW_CashPaymentAnP, -2).Value = rs.Fields("CashPaymentAnP").Value
91            .Offset(ROW_BonusStockAnP, -2).Value = rs.Fields("BonusStockAnP").Value
92            .Offset(ROW_PromoFundAnP, -2).Value = rs.Fields("PromoFundAnP").Value
93            .Offset(ROW_StaffIncentivesAnP, -2).Value = rs.Fields("StaffIncentivesAnP").Value
94            .Offset(ROW_PRAHospitalityAnP, -2).Value = rs.Fields("PRAHospitalityAnP").Value
          
              ' Reciprocal Spend
95            .Offset(ROW_ReciprocalSpend, -2).Value = rs.Fields("ReciprocalSpend").Value
96        End With
97    End If

98    Call CloseRecordset(rs, True)

      ' Delete Rows and get Contributive Margin row to be used in PEM summary
99    With rng
100       intCM_Row = .Offset(ROW_CONTRIBMARG, 0).Row
101       Select Case frm.cboRouteToMarket.Text
              Case "Direct"
102               .Offset(ROW_COP, i).EntireRow.Delete xlShiftUp
103               .Offset(ROW_KWI, i).EntireRow.Delete xlShiftUp
104               intCM_Row = .Offset(ROW_CONTRIBMARG - 2, 0).Row
105           Case "Indirect"
106               If InStr(1, GetIN_List(LBXSelectedItems(frm.lstWholesaler, 0), vbNullString, "|"), "ALM") = 0 Then
107                   .Offset(ROW_COP, i).EntireRow.Delete xlShiftUp
108                   intCM_Row = .Offset(ROW_CONTRIBMARG - 1, 0).Row
109               End If
110       End Select
111   End With

Proc_Exit:
112   wb.Application.DisplayAlerts = True
113   PopCallStack
114   Exit Sub

Err_Handler:
115   GlobalErrHandler
116   Resume Proc_Exit
End Sub

Public Sub CreateSummaryReport(wb As Workbook, frm As frmMain)
      Dim ws As Worksheet
      Dim rng As Range
      Dim arrFamily As Variant
      Dim arrProdType As Variant
      Dim i As Integer
      Dim rs As ADODB.Recordset
      Dim qry As String
      Dim strRefNumber As String

      Dim intRowCount As Integer

      Const RNG_START = "D9"
      Const ROW_VOL = 0
      Const ROW_VOL9L = 1
      Const ROW_GSV = 3
      Const ROW_BANNERTERMS = 5
      Const ROW_STANDTERMS = 6
      Const ROW_ADDNLTERMS = 7
      Const ROW_KWI = 8
      Const ROW_COP = 9
      Const ROW_QA3 = 10
      Const ROW_COOP = 11
      Const ROW_ALLOWnDISC = 12
      Const ROW_NETSALES = 14
      Const ROW_COGSnDIST = 16
      Const ROW_CONTRIBMARG = 18
      Const ROW_AnP = 20
      Const ROW_CAAP = 22
      Const ROW_GSVperVOL = 25
      Const ROW_ALLOWnDISCperGSV = 26
      Const ROW_NETSALESperVOL = 27
      Const ROW_COGSnDISTperVOL = 28
      Const ROW_CONTRIBMARGperVOL = 29
      Const ROW_CONTRIBMARGperNETSALES = 30
      Const ROW_AnPperNETSALES = 31
      Const ROW_CAAPperVOL = 32

      ' For Product Type summary
      Const RNG_PROD_TYPE_SUMM = "AE23"


1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mViewReport|CreateSummaryReport"

3     wb.Application.DisplayAlerts = False

      ' Set Categories
4     arrFamily = Array("WINE", "SPIRITS", "OTHER ALCOHOLIC BEVERAGES")
        
      ' Copy sheet template
5     wb.Worksheets(PEM_SUMM_TEMP_SHEET).Copy After:=wb.Worksheets(wb.Worksheets.Count)
6     Set ws = wb.ActiveSheet

      ' Rename sheet
7     ws.Name = PEM_SUMM_TEMP_SHEET_RENAME

      ' Get Ref number
8     strRefNumber = frm.txtRefNumber

      ' Set header
      'ws.Range("A1").Value = strRefNumber & vbCrLf & _
      '                       frm.txtFromDate.Value & " to " & frm.txtToDate & vbCrLf & _
      '                       GetIN_List(LBXSelectedItems(frm.lstContractLevelCode), "", "|")
9     ws.Range("A1").Value = strRefNumber & vbCrLf & _
                             frm.txtFromDate.Value & " to " & frm.txtToDate & vbCrLf & _
                             frm.txtOutletOrGroupName
                             

      ' Populate table
10    Set rs = New ADODB.Recordset

11    For i = 0 To UBound(arrFamily)

12        qry = vbNullString
13        qry = qry & "SELECT Sum(A.Volume) AS Sum_Volume, Sum(A.GSV) AS Sum_GSV, Sum(A.BannerTerms) AS Sum_BannerTerms, Sum(A.StandardTerms) AS Sum_StandardTerms, Sum(A.AdditionalTerms) AS Sum_AdditionalTerms, Sum(A.KWI) AS Sum_KWI, Sum(A.COP) AS Sum_COP, Sum(A.QA3) AS Sum_QA3, Sum(A.COOP) AS Sum_COOP, Sum(A.[COGSandDist]) AS [Sum_COGSandDist], Sum(A.[A&P]) AS [Sum_A&P] " & _
                      "FROM (" & _
                          "SELECT DISTINCT T3.PRODUCT_DESCRIPTION, T2.ContractedVolume AS [Volume], T2.ContractedGSV AS [GSV], " & _
                          "T5.BannerTerms, T5.StandardTerms, T5.AdditionalTerms, T2.KWI, T2.COP, T2.QA3, " & _
                          "(T6.CashPaymentCoop + T6.BonusStockCoop + T6.PromoFundCoop + T6.StaffIncentivesCoop + T6.PRAHospitalityCoop) * (T2.ContractedGSV / T7.Total_GSV) AS COOP, " & _
                          "(T6.CashPaymentAnP + T6.BonusStockAnP + T6.PromoFundAnP + T6.StaffIncentivesAnP + T6.PRAHospitalityAnP) * (T2.ContractedGSV / T7.Total_GSV) AS [A&P], " & _
                          "(T4.COGSperLitre * T2.ContractedVolume) AS [COGSandDist] " & _
                          "FROM (((((" & OP_MAIN_TBL & " AS T1 " & _
                          "INNER JOIN " & OP_PROD_DETAILS_TBL & " AS T2 ON T1.RefNumber = T2.RefNumber) " & _
                          "INNER JOIN " & PRODUCT_MAP_TBL & " AS T3 ON T2.ProductCode = T3.PRODUCT_CODE) " & _
                          "INNER JOIN " & COGSPERLTR_MAP_TBL & " AS T4 ON T3.PRODUCT_CODE = T4.ProductCode) " & _
                          "INNER JOIN " & OP_TRADING_TERMS_TBL & " AS T5 ON T2.RefNumber = T5.RefNumber AND T2.ProductCode = T5.ProductCode) " & _
                          "INNER JOIN " & OP_COOP_ANP_TBL & " AS T6 ON T2.RefNumber = T6.RefNumber) " & _
                          "INNER JOIN (SELECT RefNumber, SUM(ContractedGSV) AS Total_GSV FROM " & OP_PROD_DETAILS_TBL & " GROUP BY RefNumber) AS T7 ON T2.RefNumber = T7.RefNumber " & _
                          "WHERE T1.RefNumber = '" & strRefNumber & "' " & _
                            "AND T2.Family = '" & arrFamily(i) & "' " & _
                            "AND T4.Start_Date <=#" & GetPromoDate(End_Date, frm) & "# AND T4.End_Date >=#" & GetPromoDate(Start_Date, frm) & "# " & _
                      ") AS A"
14        rs.Open qry, cn
          
          ' Generate report
15        If Not rs.EOF Then
              ' Set starting range
16            Set rng = ws.Range(RNG_START)
              
              ' Print data
17            With rng
                  ' Volume
18                .Offset(ROW_VOL, i).Value = rs.Fields("Sum_Volume").Value
                  ' Volume 9L cases = Volume / 9
19                .Offset(ROW_VOL9L, i).Formula = "=" & GetRangeAddress(.Offset(ROW_VOL, i)) & "/9"
                  ' Gross Sales
20                .Offset(ROW_GSV, i).Value = rs.Fields("Sum_GSV").Value
                  ' Banner Terms
21                .Offset(ROW_BANNERTERMS, i).Value = -1 * rs.Fields("Sum_BannerTerms").Value
                  ' Standard Terms
22                .Offset(ROW_STANDTERMS, i).Value = -1 * rs.Fields("Sum_StandardTerms").Value
                  ' Additional Terms
23                .Offset(ROW_ADDNLTERMS, i).Value = -1 * rs.Fields("Sum_AdditionalTerms").Value
                  ' KWI
24                .Offset(ROW_KWI, i).Value = -1 * rs.Fields("Sum_KWI").Value
                  ' COP
25                .Offset(ROW_COP, i).Value = -1 * rs.Fields("Sum_COP").Value
                  ' QA3
26                .Offset(ROW_QA3, i).Value = -1 * rs.Fields("Sum_QA3").Value
                  ' COOP
27                .Offset(ROW_COOP, i).Value = -1 * rs.Fields("Sum_COOP").Value
                  ' Allowance and Discount
28                .Offset(ROW_ALLOWnDISC, i).Formula = "=SUM(" & GetRangeAddress(.Offset(ROW_BANNERTERMS, i)) & ":" & GetRangeAddress(.Offset(ROW_COOP, i)) & ")"
                  ' Net Sales = GSV + Allowance and Discount
29                .Offset(ROW_NETSALES, i).Formula = "=" & GetRangeAddress(.Offset(ROW_GSV, i)) & "-" & GetRangeAddress(.Offset(ROW_ALLOWnDISC, i))
                  ' COGS and Dist
30                .Offset(ROW_COGSnDIST, i).Value = -1 * rs.Fields("Sum_COGSandDist").Value
                  ' Contributive Margin = Net Sales + COGS and Dist
31                .Offset(ROW_CONTRIBMARG, i).Formula = "=" & GetRangeAddress(.Offset(ROW_NETSALES, i)) & "+" & GetRangeAddress(.Offset(ROW_COGSnDIST, i))
                  ' A&P
32                .Offset(ROW_AnP, i).Value = 1 * rs.Fields("Sum_A&P").Value
                  ' CAAP = Contributive Margin + A&P
33                .Offset(ROW_CAAP, i).Formula = "=" & GetRangeAddress(.Offset(ROW_CONTRIBMARG, i)) & "-" & GetRangeAddress(.Offset(ROW_AnP, i))
                  
                  ' Gross Sales/L = Gross Sales / Volume
34                .Offset(ROW_GSVperVOL, i).Formula = "=IFERROR(" & GetRangeAddress(.Offset(ROW_GSV, i)) & "/" & GetRangeAddress(.Offset(ROW_VOL, i)) & ", 0)"
                  ' Allowance&Discount % Gross Sales = Allowance and Discount / Gross Sales)
35                .Offset(ROW_ALLOWnDISCperGSV, i).Formula = "=IFERROR(-" & GetRangeAddress(.Offset(ROW_ALLOWnDISC, i)) & "/" & GetRangeAddress(.Offset(ROW_GSV, i)) & ", 0)"
                  ' Net Sales/L = Net Sales / Volume
36                .Offset(ROW_NETSALESperVOL, i).Formula = "=IFERROR(" & GetRangeAddress(.Offset(ROW_NETSALES, i)) & "/" & GetRangeAddress(.Offset(ROW_VOL, i)) & ", 0)"
                  ' COGS&Dist/L = (COGS and Dist / Volume)
37                .Offset(ROW_COGSnDISTperVOL, i).Formula = "=IFERROR(-" & GetRangeAddress(.Offset(ROW_COGSnDIST, i)) & "/" & GetRangeAddress(.Offset(ROW_VOL, i)) & ", 0)"
                  ' Contributive Margin/L = Contributive Margin / Volume
38                .Offset(ROW_CONTRIBMARGperVOL, i).Formula = "=IFERROR(" & GetRangeAddress(.Offset(ROW_CONTRIBMARG, i)) & "/" & GetRangeAddress(.Offset(ROW_VOL, i)) & ", 0)"
                  ' Contributive Margin/L = Contributive Margin / Net Sales
39                .Offset(ROW_CONTRIBMARGperNETSALES, i).Formula = "=IFERROR(" & GetRangeAddress(.Offset(ROW_CONTRIBMARG, i)) & "/" & GetRangeAddress(.Offset(ROW_NETSALES, i)) & ", 0)"
                  ' Total A&P % Net Sales = (A&P / Net Sales)
40                .Offset(ROW_AnPperNETSALES, i).Formula = "=IFERROR(-" & GetRangeAddress(.Offset(ROW_AnP, i)) & "/" & GetRangeAddress(.Offset(ROW_NETSALES, i)) & ", 0)"
                  ' CAAP/L = CAAP / Volume
41                .Offset(ROW_CAAPperVOL, i).Formula = "=IFERROR(" & GetRangeAddress(.Offset(ROW_CAAP, i)) & "/" & GetRangeAddress(.Offset(ROW_VOL, i)) & ", 0)"
42            End With 'rng
43        End If
          
44        Call CloseRecordset(rs)
45    Next i
        
46    Call CloseRecordset(rs, True)

      ' Generate formulas for Product Type summary
47    Set rng = ws.Range(RNG_PROD_TYPE_SUMM)
48    arrProdType = Split(GetItemFromMappingTbl(SETTINGS_TBL, "Settings_Value", "Settings_Name", "Product_Types", """"), "|")
49    For i = 0 To UBound(arrProdType)
50        With rng
51            .Offset(i, 0).Value = arrProdType(i)
52            .Offset(i, 1).Formula = "=SUMIFS('" & PEM_TEMP_SHEET_RENAME & "'!$D$" & intCM_Row & ":$XFD$" & intCM_Row & ",'" & PEM_TEMP_SHEET_RENAME & "'!$D$7:$XFD$7," & GetRangeAddress(.Offset(i, 0)) & ")"
53        End With
54    Next i

      ' Delete Rows
55    Set rng = ws.Range(RNG_START)
56    With rng
57        Select Case frm.cboRouteToMarket.Text
              Case "Direct"
58                .Offset(ROW_COP, i).EntireRow.Delete xlShiftUp
59                .Offset(ROW_KWI, i).EntireRow.Delete xlShiftUp
60            Case "Indirect"
61                If InStr(1, GetIN_List(LBXSelectedItems(frm.lstWholesaler, 0), vbNullString, "|"), "ALM") = 0 Then
62                    .Offset(ROW_COP, i).EntireRow.Delete xlShiftUp
63                End If
64        End Select
65    End With

      ' Calculate formulas
66    Application.Calculate

Proc_Exit:
67    wb.Application.DisplayAlerts = True
68    PopCallStack
69    Exit Sub

Err_Handler:
70    GlobalErrHandler
71    Resume Proc_Exit
End Sub

Public Sub CreateE1UploadSheet(wb As Workbook, frm As frmMain)
      Dim ws As Worksheet
      Dim rs As ADODB.Recordset
      Dim qry As String
      Dim strOutletNum As String
      Dim strOutletName As String
      Dim intRowCount As Integer
      Dim intColCount As Integer
      Dim arr As Variant
      Dim i As Integer
      Dim blnHasOutletNum As Boolean
          
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mViewReport|CreateE1UploadSheet"

      Const RNG_START = "A2"
      Const ADJ_NAME = "QA3S"
      Const THRESHOLD_UM = "CA"
      Const B_C = 5
      Const BASIS_CODE = vbNullString
      Const CUR_COD = "AUD"

3     wb.Application.DisplayAlerts = False
        
      ' Copy sheet template
4     wb.Worksheets(E1_UPLOAD_TEMP_SHEET).Copy After:=wb.Worksheets(wb.Worksheets.Count)
5     Set ws = wb.ActiveSheet

      ' Rename sheet
6     ws.Name = E1_UPLOAD_TEMP_SHEET_RENAME

7     Set rs = New ADODB.Recordset

      ' Get Outlet number
      'strOutletNum = frm.lstOutlet.List(LBXSelectedIndexes(frm.lstOutlet)(0), 1)
      'strOutletName

      '    ' ? No OUTLET_ACCOUNT_CODE
      '    ' Check if there are E1 Outlet numbers
      '    arr = LBXSelectedItems(frm.lstContractLevelCode, 2)
      '    blnHasOutletNum = False
      '    For i = 0 To UBound(arr)
      '        If Len(arr(i)) <> 0 Then
      '            blnHasOutletNum = True
      '            Exit For
      '        End If
      '    Next i

8     qry = "SELECT DISTINCT T4.AdjNamePrefix & '" & ADJ_NAME & "', '" & THRESHOLD_UM & "', Format((T2.QA3PerCaseUser + T2.QA3PerCaseAuto) * -1, '0.0000######'), " & B_C & ", '" & BASIS_CODE & "' , '', T3.OutletName, '" & CUR_COD & "', Format(T1.FromDate, 'dd/mm/yy'), Format(T1.ToDate, 'dd/mm/yy'), T2.ProductCode " & _
            "FROM ((" & OP_MAIN_TBL & " AS T1 INNER JOIN " & OP_PROD_DETAILS_TBL & " AS T2 ON T1.RefNumber = T2.RefNumber)" & IIf(blnHasOutletNum, ", " & CUSTOMER_MAP_TBL & " AS T3", " ") & " " & _
            "INNER JOIN " & OUTLET_INFO_TBL & " AS T3 ON T1.RefNumber = T3.RefNumber) " & _
            "INNER JOIN " & ADJNAME_PREFIX_MAP_TBL & " AS T4 ON T1.State = T4.State " & _
            "WHERE T1.RefNumber = '" & frm.txtRefNumber & "' " & _
            "ORDER BY T3.OutletName, T2.ProductCode "
9     rs.Open qry, cn

10    If Not rs.EOF Then
          ' Get row count
11        rs.MoveFirst
12        intRowCount = UBound(rs.GetRows(1000), 2)
          ' Get column count
13        intColCount = rs.Fields.Count
          
14        rs.MoveFirst
15        ws.Range(RNG_START).CopyFromRecordset rs
          
          ' Format table
          ' Copy first row
16        ws.Range(ws.Range(RNG_START), ws.Range(RNG_START).Offset(0, intColCount - 1)).Copy
          ' Then paste special - Formats
17        ws.Range(ws.Range(RNG_START), ws.Range(RNG_START).Offset(intRowCount, intColCount - 1)).PasteSpecial xlPasteFormats
18    End If
19    Call CloseRecordset(rs)
20    Set rs = Nothing

21    wb.Application.DisplayAlerts = True

Proc_Exit:
22    PopCallStack
23    Exit Sub

Err_Handler:
24    GlobalErrHandler
25    Resume Proc_Exit
End Sub

