Attribute VB_Name = "mDealSheets"
Option Explicit

Public Sub CreateALMDealSheet(wb As Workbook, frm As frmMain)
      Dim ws As Worksheet
      Dim strFile As String
      Dim strDealLevel As String
      Dim rs As ADODB.Recordset
      Dim qry As String
      Dim intRowCount As Integer
      Dim intColCount As Integer
      Dim arr As Variant
      Dim i As Integer
      Dim arrProducts As Variant
      Dim arrQA3 As Variant

      Const rngRefNum = "F2"
      Const rngPromoStartDate = "F5"
      Const rngPromoEndDate = "G5"
      Const rngBuyingPeriodStartDate = "D8"
      Const rngBuyingPeriodEndDate = "D9"
      Const rngCustName = "G8"
      Const rngCustNumber = "G9"
      Const rngSubmittedBy = "P10"
      Const rngReferenceCode = "P12"
      Const rngState_NSW = "C12"
      Const rngState_VIC = "C13"
      Const rngState_QLD = "C14"
      Const rngState_SA = "C15"
      Const rngState_WA = "E12"
      Const rngState_TAS = "E13"
      Const rngState_NT = "E14"
      Const rngState_ACT = "E15"
      Const rngGroupName = "G13"
      Const rngQtyRestrictionNumCases = "S15"
      Const rngProductInfo = "E25"
      Const rngInsertPoint = "A45"


1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mDealSheets|CreateALMDealSheet"

3     wb.Application.DisplayAlerts = False

      ' Copy sheet template
4     wb.Worksheets(ALM_DEAL_TEMP_SHEET).Copy After:=wb.Worksheets(wb.Worksheets.Count)
5     Set ws = wb.ActiveSheet

      ' Rename sheet
6     ws.Name = ALM_DEAL_TEMP_SHEET_RENAME

      ' Populate the deal sheet
7     ws.Range(rngRefNum).Value = ""
8     ws.Range(rngPromoStartDate).Value = ""
9     ws.Range(rngPromoEndDate).Value = ""
10    ws.Range(rngBuyingPeriodStartDate).Value = frm.txtFromDate 'Format(DateAdd("ww", -2, frm.txtFromDate), "dd-mmm-yyyy")
11    ws.Range(rngBuyingPeriodEndDate).Value = frm.txtToDate 'Format(DateAdd("ww", 1, frm.txtToDate), "dd-mmm-yyyy")

      'Select Case frm.cboContractLevel.Value
      '    Case "OP Outlet Level"
      '        ws.Range(rngCustName).Value = frm.txtOutletOrGroupName
      '    Case "OP Banner"
      '        ws.Range(rngState_NSW).Value = "*"
      '        ws.Range(rngState_VIC).Value = "*"
      '        ws.Range(rngState_QLD).Value = "*"
      '        ws.Range(rngState_SA).Value = "*"
      '        ws.Range(rngState_WA).Value = "*"
      '        ws.Range(rngState_TAS).Value = "*"
      '        ws.Range(rngState_NT).Value = "*"
      '        ws.Range(rngState_ACT).Value = "*"
      '    Case "OP Banner Region"
      '        Select Case GetItemFromMappingTbl(CUSTOMER_MAP_TBL, "State", "BannerRegionCode", GetIN_List(LBXSelectedItems(frm.lstContractLevelCode, 1), vbNullString, "|"), "'")
      '            Case "NAT"
      '                ws.Range(rngState_NSW).Value = "*"
      '                ws.Range(rngState_VIC).Value = "*"
      '                ws.Range(rngState_QLD).Value = "*"
      '                ws.Range(rngState_SA).Value = "*"
      '                ws.Range(rngState_WA).Value = "*"
      '                ws.Range(rngState_TAS).Value = "*"
      '                ws.Range(rngState_NT).Value = "*"
      '                ws.Range(rngState_ACT).Value = "*"
      '            Case "NSW"
      '                ws.Range(rngState_NSW).Value = "*"
      '            Case "VIC"
      '                ws.Range(rngState_VIC).Value = "*"
      '            Case "QLD"
      '                ws.Range(rngState_QLD).Value = "*"
      '            Case "SA"
      '                ws.Range(rngState_SA).Value = "*"
      '            Case "WA"
      '                ws.Range(rngState_WA).Value = "*"
      '            Case "TAS"
      '                ws.Range(rngState_TAS).Value = "*"
      '            Case "NT"
      '                ws.Range(rngState_NT).Value = "*"
      '            Case "ACT"
      '                ws.Range(rngState_ACT).Value = "*"
      '            Case Else
      '                ' nothing
      '        End Select
      '
      '    Case Else
      '        ' nothing
      'End Select

12    ws.Range(rngSubmittedBy).Value = GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "Name", "ID", frm.cboCreator, "'")
13    ws.Range(rngReferenceCode).Value = frm.txtRefNumber

      ' Get Products and QA3 data
14    arrProducts = frm.lstProducts.List
15    arrQA3 = frm.lstQA3.List

16    If frm.lstProducts.ListCount <> 0 Then
17        For i = 0 To UBound(arrProducts)
18            With ws.Range(rngProductInfo)
19                .Offset(i, 0).Value = ""
20                .Offset(i, 1).Value = arrProducts(i, ProdList_ProdDesc)
21                .Offset(i, 3).Value = arrProducts(i, ProdList_BottleSize)
22                .Offset(i, 4).Value = arrProducts(i, ProdList_UnitsPerCase)
23                .Offset(i, 5).Value = arrQA3(i, QA3List_QA3AutoRoundoff) + arrQA3(i, QA3List_QA3InputRoundoff)
          
                  ' Add more rows if it exceeds 20 records
24                If i > 20 Then ws.Range(rngInsertPoint).EntireRow.Insert xlShiftDown
          
25            End With
26        Next i
          
          ' Format table
          ' Copy first row
27        ws.Range(ws.Range(rngProductInfo), ws.Range(rngProductInfo).Offset(0, 18)).Copy
          ' Then paste special - Formats
28        ws.Range(ws.Range(rngProductInfo), ws.Range(rngProductInfo).Offset(intRowCount, 18)).PasteSpecial xlPasteFormats
29    End If

30    wb.Application.DisplayAlerts = True

Proc_Exit:
31    PopCallStack
32    Exit Sub

Err_Handler:
33    GlobalErrHandler
34    Resume Proc_Exit
End Sub

Public Sub CreateStandardDealSheet(wb As Workbook, frm As frmMain, strWholesaler As String)
      Dim ws As Worksheet
      Dim strFile As String
      Dim strDealLevel As String
      Dim rs As ADODB.Recordset
      Dim qry As String
      Dim intRowCount As Integer
      Dim intColCount As Integer
      Dim arr As Variant
      Dim i As Integer
      Dim arrProducts As Variant
      Dim arrQA3 As Variant

      Const rngRefNum = "C3"
      Const rngPRAContactName = "C7"
      Const rngPRAContactPhone = "C8"
      Const rngPRAContactEmail = "C9"
      Const rngPRAAuthoriserName = "C10"
      Const rngPRAAuthoriserEmail = "C11"
      Const rngCustomerContactName = "C14"
      Const rngCustomerContactPhone = "C15"
      Const rngCustomerContactEmail = "C16"
      Const rngBanner = "C21"
      Const rngBannerRegionName = "C22"
      Const rngOutletName = "C23"
      Const rngOutetNumber = "C24"
      Const rngState_All = "I21"
      Const rngState_NSW_ACT = "I22"
      Const rngState_VIC_TAS = "I23"
      Const rngState_QLD = "I24"
      Const rngState_WA = "K22"
      Const rngState_SA = "K23"
      Const rngState_NT = "K24"
      Const rngPromoStartDate = "C27"
      Const rngPromoEndDate = "E27"
      Const rngBuyingPeriodStartDate = "C28"
      Const rngBuyingPeriodEndDate = "E28"
      Const rngPerCustInGrpMixBuy = "I27"
      Const rngPerBannRegMixBuy = "I28"
      Const rngPerCustInGrpNumCases = "J27"
      Const rngPerBannRegNumCases = "J28"
      Const rngComments = "M4"
      Const rngProductInfo = "B39"
          
          
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mDealSheets|CreateStandardDealSheet"

3     wb.Application.DisplayAlerts = False

      ' Copy sheet template
4     wb.Worksheets(STANDARD_DEAL_TEMP_SHEET).Copy After:=wb.Worksheets(wb.Worksheets.Count)
5     Set ws = wb.ActiveSheet
         
      ' Rename sheet to Wholesaler code
6     ws.Name = strWholesaler
          
      ' Populate the deal sheet
7     ws.Range(rngRefNum).Value = frm.txtRefNumber
8     ws.Range(rngPRAContactName).Value = GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "Name", "ID", frm.cboCreator, "'")
9     ws.Range(rngPRAContactPhone).Value = "" 'GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "Phone", "[Name]", frm.cboPromoCreator.Value, """")
10    ws.Range(rngPRAContactEmail).Value = "" 'GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "Email", "[Name]", frm.cboPromoCreator.Value, """")
11    ws.Range(rngPRAAuthoriserName).Value = GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "Name", "ID", GetItemFromMappingTbl(PRA_MANAGER_TBL, "[Name]", "ID", CStr(GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "ManagerID", "ID", frm.cboCreator.Value, """"))), "'")
12    ws.Range(rngPRAAuthoriserEmail).Value = "" 'GetItemFromMappingTbl(PRA_MANAGER_TBL, "Email", "[Name]", frm.cboPromoManager.Value, """")
13    ws.Range(rngCustomerContactName).Value = frm.txtCustomerName
14    ws.Range(rngCustomerContactPhone).Value = frm.txtCustomerPhone
15    ws.Range(rngCustomerContactEmail).Value = frm.txtCustomerEmail

16    Select Case frm.cboContractLevel.Value
          Case "OP Banner"
17            ws.Range(rngBanner).Value = frm.txtOutletOrGroupName
18        Case "OP Banner Region"
19            ws.Range(rngBannerRegionName).Value = frm.txtOutletOrGroupName
20        Case "OP Outlet Level"
21            ws.Range(rngOutletName).Value = frm.txtOutletOrGroupName
22    End Select

      '    strDealLevel = GetItemFromMappingTbl(PROMO_MAIN_TBL, "DealLevel", "RefNumber", frm.txtPromoRefNumber, "'")
      '    If strDealLevel = "Banner" Then
      '        ws.Range(rngState_All).Value = "*"
      '    Else
      '        If strDealLevel = "BannerRegionName" Then
      '            arr = Array(frm.cboBannerRegionName.Value)
      '        Else
      '            arr = Split(GetItemFromMappingTbl(PROMO_MAIN_TBL, "BannerRegionOfOutlets", "RefNumber", frm.txtPromoRefNumber, "'"), "|")
      '        End If
      '        For i = 0 To UBound(arr)
      '            Select Case UCase(Left(arr(i), InStr(1, arr(i), " ") - 1))
      '                Case "NAT"
      '                    ws.Range(rngState_All).Value = "*"
      '                Case "NSW"
      '                    ws.Range(rngState_NSW_ACT).Value = "*"
      '                Case "ACT"
      '                    ws.Range(rngState_NSW_ACT).Value = "*"
      '                Case "NT"
      '                    ws.Range(rngState_NT).Value = "*"
      '                Case "QLD"
      '                    ws.Range(rngState_QLD).Value = "*"
      '                Case "SA"
      '                    ws.Range(rngState_SA).Value = "*"
      '                Case "TAS"
      '                    ws.Range(rngState_VIC_TAS).Value = "*"
      '                Case "VIC"
      '                    ws.Range(rngState_VIC_TAS).Value = "*"
      '                Case "WA"
      '                    ws.Range(rngState_WA).Value = "*"
      '                Case Else
      '                    ' nothing
      '            End Select
      '        Next i
      '    End If

23    ws.Range(rngPromoStartDate).Value = ""
24    ws.Range(rngPromoEndDate).Value = ""
25    ws.Range(rngBuyingPeriodStartDate).Value = frm.txtFromDate
26    ws.Range(rngBuyingPeriodEndDate).Value = frm.txtToDate

27    ws.Range(rngComments).Value = frm.txtComments

      'Select Case frm.cboQtyRestriction.Value
      '    Case "Per Customer in Group"
      '        ws.Range(rngPerCustInGrpMixBuy).Value = "Y"
      '        ws.Range(rngPerCustInGrpNumCases).Value = frm.txtQtyRestrictionNumOfCases
      '    Case "Per Banner Region"
      '        ws.Range(rngPerBannRegMixBuy).Value = "Y"
      '        ws.Range(rngPerBannRegNumCases).Value = frm.txtQtyRestrictionNumOfCases
      'End Select

      ' Get Products and QA3 data
28    arrProducts = frm.lstProducts.List
29    arrQA3 = frm.lstQA3.List

30    If frm.lstProducts.ListCount <> 0 Then
31        For i = 0 To UBound(arrProducts)
32            With ws.Range(rngProductInfo)
33                .Offset(i, 0).Value = arrProducts(i, ProdList_Brand)
34                .Offset(i, 1).Value = arrProducts(i, ProdList_Subbrand)
35                .Offset(i, 2).Value = arrProducts(i, ProdList_ProdDesc)
36                .Offset(i, 3).Value = " "
37                .Offset(i, 4).Value = arrProducts(i, ProdList_BottleSize)
38                .Offset(i, 5).Value = arrProducts(i, ProdList_UnitsPerCase)
39                .Offset(i, 6).Value = arrQA3(i, QA3List_QA3AutoRoundoff) + arrQA3(i, QA3List_QA3InputRoundoff)
40            End With
41        Next i
          
42        intRowCount = UBound(arrProducts) + 1
          
          ' Format table
          ' Copy first row
43        ws.Range(ws.Range(rngProductInfo), ws.Range(rngProductInfo).Offset(0, 18)).Copy
          ' Then paste special - Formats
44        ws.Range(ws.Range(rngProductInfo), ws.Range(rngProductInfo).Offset(intRowCount, 18)).PasteSpecial xlPasteFormats
45    End If

46    wb.Application.DisplayAlerts = True

Proc_Exit:
47    PopCallStack
48    Exit Sub

Err_Handler:
49    GlobalErrHandler
50    Resume Proc_Exit
End Sub

