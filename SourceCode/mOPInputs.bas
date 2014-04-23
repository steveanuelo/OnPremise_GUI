Attribute VB_Name = "mOPInputs"
Option Explicit

Public Const OP_MAIN_TBL = "T_Main_Details"
Public Const OP_PROD_DETAILS_TBL = "T_Main_Product_Details"
Public Const OP_COOP_ANP_TBL = "T_Main_Coop_And_AnP"
Public Const OP_TRADING_TERMS_TBL = "T_Main_TradingTerms"
Public Const OP_TRADING_TERMS_CONST_TBL = "T_Main_TradingTerms_Const"
Public Const OP_PROD_NON_QA3_TBL = "T_Main_Product_Non_QA3"

Public Function ValidateInputs(frm As frmMain) As Boolean
      Dim x As Integer
      Dim dblCtr As Double
          
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mOPInputs|ValidateInputs"

3     ValidateInputs = True

      ' Ref#
4     If HasEmptyValues(frm.txtRefNumber, "OP Contract Ref#") Then GoTo Validation_Error

      ' PRA Contact Information
5     If HasEmptyValues(frm.cboCreator, "Creator") Then GoTo Validation_Error
6     If HasEmptyValues(frm.cboManager, "Manager/Authoriser") Then GoTo Validation_Error

      ' Customer Contact Information
7     If HasEmptyValues(frm.txtOutletOrGroupName, "Outlet/Group Name") Then GoTo Validation_Error
8     If HasEmptyValues(frm.txtAddress, "Address") Then GoTo Validation_Error
9     If HasEmptyValues(frm.txtCustomerName, "Customer Contact Name") Then GoTo Validation_Error
10    If HasEmptyValues(frm.txtCustomerPhone, "Customer Contact Phone Number") Or _
         HasInvalidNumber(frm.txtCustomerPhone, strName:="Customer Contact Phone Number") Then GoTo Validation_Error
11    If HasEmptyValues(frm.txtCustomerEmail, "Customer Contact Email") Then GoTo Validation_Error

      ' Contract Level
12    If HasEmptyValues(frm.cboContractLevel, "Contract Level") Then GoTo Validation_Error
13    If LBXSelectCount(frm.lstContractLevelCode) = 0 Then
14        MsgBox "Please select a Contract Level Code.", vbExclamation, "Data Validation Error"
15        GoTo Validation_Error
16    End If

      ' Contract Type
17    If Not (frm.chkSpirits.Value = True Or frm.chkChampagne.Value = True Or frm.chkWine.Value = True) Then
18        MsgBox "At least one Contract Type should be selected.", vbExclamation, "Data Validation Error"
19        GoTo Validation_Error
20    End If

      ' Contract Form
21    If HasEmptyValues(frm.cboContractForm, "Contract Form") Then GoTo Validation_Error

      ' Contract Dates
22    If HasEmptyValues(frm.txtFromDate, "From Date") Then GoTo Validation_Error
23    If HasEmptyValues(frm.txtToDate, "To Date") Then GoTo Validation_Error
24    If Not ((CDate(frm.txtFromDate) < CDate(frm.txtToDate)) And (CDate(frm.txtToDate) > CDate(frm.txtFromDate))) Then
25        MsgBox "Invalid period dates", vbExclamation, "Data Validation Error"
26        GoTo Validation_Error
27    End If
28    If Len(frm.txtFromExtnDate) <> 0 Then
29        If HasEmptyValues(frm.txtToExtnDate, "To Date Extension") Then GoTo Validation_Error
30    End If
31    If Len(frm.txtToExtnDate) <> 0 Then
32        If HasEmptyValues(frm.txtFromExtnDate, "From Date Extension") Then GoTo Validation_Error
33    End If
34    If Len(frm.txtFromExtnDate) <> 0 And Len(frm.txtToExtnDate) <> 0 Then
35        If Not ((CDate(frm.txtFromExtnDate) < CDate(frm.txtToExtnDate)) And (CDate(frm.txtToExtnDate) > CDate(frm.txtFromExtnDate))) Then
36            MsgBox "Invalid extension period dates", vbExclamation, "Data Validation Error"
37            GoTo Validation_Error
38        End If
39    End If


      ' Route to Market
40    If HasEmptyValues(frm.cboRouteToMarket, "Route to Market") Then
41        GoTo Validation_Error
42    Else
43        If frm.cboRouteToMarket <> "Direct" And LBXSelectCount(frm.lstWholesaler) = 0 Then
44            MsgBox "Please select a Wholesaler.", vbExclamation, "Data Validation Error"
45            GoTo Validation_Error
46        End If
47    End If

      ' Products list
48    If frm.lstProducts.ListCount = 0 Then
49        MsgBox "Please fill up the Products tab.", vbExclamation, "Data Validation Error"
50        GoTo Validation_Error
51    End If

      ' Terms
52    If HasInvalidNumber(frm.txtAllProd_PctGSV, True, "All Products % GSV-QA3") Then GoTo Validation_Error
53    If HasInvalidNumber(frm.txtAllProd_DollarPerLitre, True, "All Products $ Per Litre") Then GoTo Validation_Error
54    If HasInvalidNumber(frm.txtAllNonContract_PctGSV, True, "All Non Contracted Products % GSV-QA3") Then GoTo Validation_Error
55    If HasInvalidNumber(frm.txtAllNonContract_DollarPerLitre, True, "All Non Contracted Products $ Per Litre") Then GoTo Validation_Error
56    If HasInvalidNumber(frm.txtTTBannerGSV, True, "Banner Terms % GSV") Then GoTo Validation_Error
57    If HasInvalidNumber(frm.txtTTBannerGSVlessQA3, True, "Banner Terms % GSV-QA3") Then GoTo Validation_Error

      ' COOP and AnP
58    If Len(frm.txtTotalCashPay) <> 0 Then _
          If HasEmptyValues(frm.txtCommentsCashPay, "Cash Payments Comments") Then GoTo Validation_Error
59    If Len(frm.txtTotalBonusStock) <> 0 Then _
          If HasEmptyValues(frm.txtCommentsBonusStock, "Bonus Stock Comments") Then GoTo Validation_Error
60    If Len(frm.txtTotalPromoFund) <> 0 Then _
          If HasEmptyValues(frm.txtCommentsPromoFund, "Promotional Fund Comments") Then GoTo Validation_Error
61    If Len(frm.txtTotalStaffIncentives) <> 0 Then _
          If HasEmptyValues(frm.txtCommentsStaffIncentives, "Staff Incentives Comments") Then GoTo Validation_Error
62    If Len(frm.txtTotalPRAHospitality) <> 0 Then _
          If HasEmptyValues(frm.txtCommentsPRAHospitality, "PRA Hospitality Comments") Then GoTo Validation_Error
63    If Len(frm.txtReciprocalSpend) <> 0 Then _
          If HasEmptyValues(frm.txtReciprocalSpendComments, "Reciprocal Spend Comments") Then GoTo Validation_Error

Proc_Exit:
64    PopCallStack
65    Exit Function

Err_Handler:
66    GlobalErrHandler

Validation_Error:
67    ValidateInputs = False
68    GoTo Proc_Exit
End Function

Public Function SaveRecordToDB(frm As frmMain, strSaveType As String) As Integer
      Dim strRefNumber As String
      Dim qry As String
      Dim x As Integer, y As Integer
      Dim rng As Range
      Dim arrData As Variant
      Dim arrData1 As Variant
      Dim dblTotalGSV As Double


1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mOPInputs|SaveRecordToDB"

      ' Get Reference number
3     strRefNumber = frm.txtRefNumber

      ' Check if Ref# is already existing.
      ' If yes update, else add a new record
4     If IsItemExistInTable(OP_MAIN_TBL, "RefNumber", strRefNumber, "'") Then
          ' Update
          
          ' Main table
5         qry = vbNullString
6         qry = qry & "UPDATE " & OP_MAIN_TBL & " " & _
                      "SET CreatorID = """ & frm.cboCreator.Value & """, " & _
                          "OutletOrGroupName = """ & frm.txtOutletOrGroupName & """, " & _
                          "CustomerAddress = """ & frm.txtAddress & """, " & _
                          "CustomerName = """ & frm.txtCustomerName & """, " & _
                          "CustomerPhone = " & SetEmptyValue(frm.txtCustomerPhone, NullForDB) & ", " & _
                          "CustomerEmail = """ & frm.txtCustomerEmail & """, " & _
                          "ContractLevel = """ & frm.cboContractLevel.Value & """, "
                            
          ' Check if there are Contract Level Code selected
7         If LBXSelectCount(frm.lstContractLevelCode) <> 0 Then
8             qry = qry & "ContractLevelCode = """ & GetIN_List(LBXSelectedItems(frm.lstContractLevelCode, 1), vbNullString, "|") & """, "
9         Else
10            qry = qry & "ContractLevelCode = '', "
11        End If

          ' State
12        Select Case frm.cboContractLevel.Value
              Case "OP Banner"
13                qry = qry & "State = '" & GetItemFromMappingTbl(CUSTOMER_MAP_TBL, "State", "BannerCode", CStr(LBXSelectedItems(frm.lstContractLevelCode, 1)(0)), "'") & "', "
14            Case "OP Banner Region"
15                qry = qry & "State = '" & GetItemFromMappingTbl(CUSTOMER_MAP_TBL, "State", "BannerRegionCode", CStr(LBXSelectedItems(frm.lstContractLevelCode, 1)(0)), "'") & "', "
16            Case "OP Outlet Level"
17                qry = qry & "State = '" & GetItemFromMappingTbl(CUSTOMER_MAP_TBL, "State", "ExternalID", CStr(LBXSelectedItems(frm.lstContractLevelCode, 1)(0)), "'") & "', "
18        End Select

19        qry = qry & "ContractType = '" & Left(IIf(frm.chkSpirits.Value = True, "Spirits|", "") & _
                                                IIf(frm.chkChampagne.Value = True, "Champagne|", "") & _
                                                IIf(frm.chkWine.Value = True, "Wine|", ""), _
                                           Len(IIf(frm.chkSpirits.Value = True, "Spirits|", "") & _
                                               IIf(frm.chkChampagne.Value = True, "Champagne|", "") & _
                                               IIf(frm.chkWine.Value = True, "Wine|", "")) - 1) & "', " & _
                      "ContractForm = """ & frm.cboContractForm.Value & """, " & _
                      "VarPrice = """ & frm.cboVarPrice.Value & """, " & _
                      "PROS = """ & frm.cboPROS.Value & """, " & _
                      "ToDate = " & IIf(Len(frm.txtToDate) <> 0, "#" & frm.txtToDate & "#", "NULL") & ", " & _
                      "FromDate = " & IIf(Len(frm.txtFromDate) <> 0, "#" & frm.txtFromDate & "#", "NULL") & ", " & _
                      "ToDate_Extention = " & IIf(Len(frm.txtToExtnDate) <> 0, "#" & frm.txtToExtnDate & "#", "NULL") & ", " & _
                      "FromDate_Extention = " & IIf(Len(frm.txtFromExtnDate) <> 0, "#" & frm.txtFromExtnDate & "#", "NULL") & ", " & _
                      "RouteToMarket = """ & frm.cboRouteToMarket.Value & """, "
                      
          ' Check if there are Wholesalers selected
20        If IsArrayAllocated(LBXSelectedItems(frm.lstWholesaler)) Then
21            qry = qry & "Wholesaler = """ & GetIN_List(LBXSelectedItems(frm.lstWholesaler), vbNullString, "|") & """, "
22        Else
23            qry = qry & "Wholesaler = '', "
24        End If
          
25        qry = qry & "Comments = """ & frm.txtComments & """, " & _
                      "LastSyncDate = #" & ConvertLocalToGMT(Now) & "# "
          
          ' Update Submission date if only saved as "For Approval"
26        If strSaveType = "For Approval" Then
27            qry = qry & ", SubmitDate = #" & Format(Now, "dd-mmm-yyyy") & "# "
28        ElseIf strSaveType = "View" Then
              'qry = qry & ", SubmitDate = #" & GetItemFromMappingTbl(OP_MAIN_TBL, "SubmitDate", "RefNumber", frm.txtRefNumber, "'") & "# "
29            qry = qry & ", SubmitDate = NULL "
30        Else ' If saving as "Draft" set Submission date as null
31            qry = qry & ", SubmitDate = NULL "
32        End If

          ' Do not update StatusID if viewing only
33        If strSaveType <> "View" Then
34            qry = qry & ", StatusID = " & GetItemFromMappingTbl(STATUS_TBL, "ID", "Description", strSaveType, "'") & " "
35        End If
                      
36        qry = qry & "WHERE RefNumber='" & strRefNumber & "';"
          
37        cn.Execute qry
          
38    Else
          ' Append
          
          ' Main table
39        qry = vbNullString
40        qry = qry & "INSERT INTO " & OP_MAIN_TBL & "(RefNumber, CreatorID, OutletOrGroupName, CustomerAddress, customerName, CustomerPhone, customerEmail, ContractLevel, " & _
                      "ContractLevelCode, State, ContractType, ContractForm, VarPrice, PROS, FromDate, ToDate, FromDate_Extention, ToDate_Extention," & _
                      "RouteToMarket, Wholesaler, Comments, LastSyncDate "
          
          ' Update Submission date if only saved as "For Approval"
41        If strSaveType = "For Approval" Then qry = qry & ", SubmitDate "
          
          ' Do not update StatusID if viewing only
42        If strSaveType <> "View" Then qry = qry & ", StatusID "
                      
43        qry = qry & ") VALUES ( " & _
                      "'" & strRefNumber & "', " & _
                      """" & frm.cboCreator.Value & """, " & _
                      """" & frm.txtOutletOrGroupName & """, " & _
                      """" & frm.txtAddress & """, " & _
                      """" & frm.txtCustomerName & """, " & _
                      SetEmptyValue(frm.txtCustomerPhone, NullForDB) & ", " & _
                      """" & frm.txtCustomerEmail & """, " & _
                      """" & frm.cboContractLevel.Value & """, "
          
          ' Check if there are Contract Level Code selected
44        If LBXSelectCount(frm.lstContractLevelCode) <> 0 Then
45            qry = qry & """" & GetIN_List(LBXSelectedItems(frm.lstContractLevelCode, 1), vbNullString, "|") & """, "
46        Else
47            qry = qry & "'', "
48        End If
          
          ' State
49        Select Case frm.cboContractLevel.Value
              Case "OP Banner"
50                qry = qry & """" & GetItemFromMappingTbl(CUSTOMER_MAP_TBL, "State", "BannerCode", CStr(LBXSelectedItems(frm.lstContractLevelCode, 1)(0)), "'") & """, "
51            Case "OP Banner Region"
52                qry = qry & """" & GetItemFromMappingTbl(CUSTOMER_MAP_TBL, "State", "BannerRegionCode", CStr(LBXSelectedItems(frm.lstContractLevelCode, 1)(0)), "'") & """, "
53            Case "OP Outlet Level"
54                qry = qry & """" & GetItemFromMappingTbl(CUSTOMER_MAP_TBL, "State", "ExternalID", CStr(LBXSelectedItems(frm.lstContractLevelCode, 1)(0)), "'") & """, "
55        End Select
          
56        qry = qry & """" & Left(IIf(frm.chkSpirits.Value = True, "Spirits|", "") & _
                                  IIf(frm.chkChampagne.Value = True, "Champagne|", "") & _
                                  IIf(frm.chkWine.Value = True, "Wine|", ""), _
                             Len(IIf(frm.chkSpirits.Value = True, "Spirits|", "") & _
                                 IIf(frm.chkChampagne.Value = True, "Champagne|", "") & _
                                 IIf(frm.chkWine.Value = True, "Wine|", "")) - 1) & """, " & _
                      """" & frm.cboContractForm.Value & """, " & _
                      """" & frm.cboVarPrice.Value & """, " & _
                      """" & frm.cboPROS.Value & """, " & _
                      IIf(Len(frm.txtFromDate) <> 0, "#" & frm.txtFromDate & "#", "NULL") & ", " & _
                      IIf(Len(frm.txtToDate) <> 0, "#" & frm.txtToDate & "#", "NULL") & ", " & _
                      IIf(Len(frm.txtFromExtnDate) <> 0, "#" & frm.txtFromExtnDate & "#", "NULL") & ", " & _
                      IIf(Len(frm.txtToExtnDate) <> 0, "#" & frm.txtToExtnDate & "#", "NULL") & ", " & _
                      """" & frm.cboRouteToMarket.Value & """, "

          ' Check if there are Wholesalers selected
57        If IsArrayAllocated(LBXSelectedItems(frm.lstWholesaler)) Then
58            qry = qry & """" & GetIN_List(LBXSelectedItems(frm.lstWholesaler), vbNullString, "|") & """, "
59        Else
60            qry = qry & "'', "
61        End If

62        qry = qry & """" & frm.txtComments & """, " & _
                      "#" & ConvertLocalToGMT(Now) & "# "

          ' Update Submission date if only saved as "For Approval"
63        If strSaveType = "For Approval" Then qry = qry & ",#" & Format(Now, "dd-mmm-yyyy") & "# "

          ' Do not update StatusID if viewing only
64        If strSaveType <> "View" Then
65            qry = qry & ", " & GetItemFromMappingTbl(STATUS_TBL, "ID", "Description", strSaveType, "'")
66        End If
                      
67        qry = qry & ")"
          
68        cn.Execute qry
          
69    End If

      ' Refresh Outlet Info
      ' Delete all records with the Ref#
70    qry = vbNullString
71    qry = "DELETE * " & _
            "FROM " & OUTLET_INFO_TBL & " " & _
            "WHERE RefNumber = '" & strRefNumber & "';"
72    cn.Execute qry

73    If frm.cboContractLevel.Value = "OP Outlet Level" Then
74        For x = 0 To LBXSelectCount(frm.lstContractLevelCode) - 1
              ' Append
75            qry = vbNullString
76            qry = qry & "INSERT INTO " & OUTLET_INFO_TBL & "(RefNumber, ExternalID, MatchCode, OutletName) " & _
                          "VALUES ( " & _
                          """" & strRefNumber & """, " & _
                          """" & LBXSelectedItems(frm.lstContractLevelCode, 1)(x) & """, " & _
                          """" & LBXSelectedItems(frm.lstContractLevelCode, 2)(x) & """, " & _
                          """" & LBXSelectedItems(frm.lstContractLevelCode, 0)(x) & """);"
77            cn.Execute qry
78        Next x
79    End If

      ' Refresh Product details
      ' Delete all records with the Ref#
80    qry = "DELETE * " & _
            "FROM " & OP_PROD_DETAILS_TBL & " " & _
            "WHERE RefNumber = '" & strRefNumber & "';"
81    cn.Execute qry

      ' Store data in array
82    arrData = frm.lstProducts.List
83    arrData1 = frm.lstQA3.List

      ' Get total GSV
84    dblTotalGSV = 0
85    For x = 0 To UBound(arrData)
86        dblTotalGSV = dblTotalGSV + CDbl(arrData(x, ProdList_ContractGSV))
87    Next x

      ' Loop for each record
88    For x = 0 To UBound(arrData)
89        qry = "INSERT INTO " & OP_PROD_DETAILS_TBL & "(RefNumber, ProductType, BrandCode, SubBrandCode, ProductCode, BottleSize, UnitsPerCase, ContractedCases, ContractedVolume, ContractedGSV, DirectPrice, WholesalePrice, QA3PerCaseUser, NIPOrLUCAuto, NIPOrLUCUser, QA3PerCaseAuto, QA3, KWI, COP, COOP, AnP, COGSnDistr, Family) " & _
                "VALUES ( " & _
                """" & strRefNumber & """, " & _
                """" & arrData(x, ProdList_ProdType) & """, " & _
                """" & arrData(x, ProdList_BrandCode) & """, " & _
                """" & arrData(x, ProdList_SubbrandCode) & """, " & _
                """" & arrData(x, ProdList_ProdCode) & """, " & _
                SetEmptyValue(arrData(x, ProdList_BottleSize), ZeroValue) & ", " & _
                SetEmptyValue(arrData(x, ProdList_UnitsPerCase), ZeroValue) & ", " & _
                SetEmptyValue(arrData(x, ProdList_ContractCases), ZeroValue) & ", " & _
                SetEmptyValue(arrData(x, ProdList_ContractVol), ZeroValue) & ", " & _
                SetEmptyValue(arrData(x, ProdList_ContractGSV), ZeroValue) & ", " & _
                SetEmptyValue(arrData1(x, QA3List_DirectPrice), ZeroValue) & ", " & _
                SetEmptyValue(arrData1(x, QA3List_WSPrice), ZeroValue) & ", " & _
                SetEmptyValue(arrData1(x, QA3List_QA3Input), ZeroValue) & ", " & _
                SetEmptyValue(arrData1(x, QA3List_NipOrLUCAuto), ZeroValue) & ", " & _
                SetEmptyValue(arrData1(x, QA3List_NipOrLUCInput), ZeroValue) & ", " & _
                SetEmptyValue(arrData1(x, QA3List_QA3Auto), ZeroValue) & ", " & _
                SetEmptyValue(arrData1(x, QA3List_QA3), ZeroValue) & ", " & _
                SetEmptyValue(arrData1(x, QA3List_KWI), ZeroValue) & ", " & _
                SetEmptyValue(arrData1(x, QA3List_COP), ZeroValue) & ", " & _
                SetEmptyValue((ConvToDbl(frm.txtCoopCashPay) + ConvToDbl(frm.txtCoopBonusStock) + ConvToDbl(frm.txtCoopPromoFund) + ConvToDbl(frm.txtCoopStaffIncentives) + ConvToDbl(frm.txtCoopPRAHospitality)) * (SetEmptyValue(arrData(x, ProdList_ContractGSV), ZeroValue) / dblTotalGSV), ZeroValue) & ", " & _
                SetEmptyValue((ConvToDbl(frm.txtAnPCashPay) + ConvToDbl(frm.txtAnPBonusStock) + ConvToDbl(frm.txtAnPPromoFund) + ConvToDbl(frm.txtAnPStaffIncentives) + ConvToDbl(frm.txtAnPPRAHospitality)) * (SetEmptyValue(arrData(x, ProdList_ContractGSV), ZeroValue) / dblTotalGSV), ZeroValue) & ", " & _
                SetEmptyValue(SetEmptyValue(arrData(x, ProdList_ContractVol), ZeroValue) * GetItemFromMappingTbl(COGSPERLTR_MAP_TBL, "COGSperLitre", strWhereCondit:="ProductCode = """ & CStr(arrData(x, ProdList_ProdCode)) & """ AND Start_Date <=#" & GetPromoDate(End_Date, frm) & "# AND End_Date >=#" & GetPromoDate(Start_Date, frm) & "#"), ZeroValue) & ", " & _
                """" & SetEmptyValue(arrData1(x, QA3List_Family), ZeroValue) & """);"
90        cn.Execute qry
91    Next x


      ' Refresh Trading Terms Input
      ' Delete all records with the Ref#
92    qry = vbNullString
93    qry = "DELETE * " & _
            "FROM " & OP_TRADING_TERMS_TBL & " " & _
            "WHERE RefNumber = '" & strRefNumber & "';"
94    cn.Execute qry

      ' Store data in array
95    arrData = frm.lstTrdTerms.List

      ' Loop for each record
96    For x = 0 To UBound(arrData)
97        qry = "INSERT INTO " & OP_TRADING_TERMS_TBL & "(RefNumber, ProductCode, DollarPerLiter, PctOfGSV, FreqOfPayments, AddnlDollarPerLiter, AddnlPctOfGSV, CondTermComments, StandardTerms, AdditionalTerms, BannerTerms) " & _
                "VALUES ( " & _
                """" & strRefNumber & """, " & _
                """" & arrData(x, TTList_ProdCode) & """, " & _
                SetEmptyValue(arrData(x, TTList_TTLtr), ZeroValue) & ", " & _
                SetEmptyValue(arrData(x, TTList_TTGSV), ZeroValue) & ", " & _
                """" & arrData(x, TTList_FreqOfPayment) & """, " & _
                SetEmptyValue(arrData(x, TTList_TTMaxLtr), ZeroValue) & ", " & _
                SetEmptyValue(arrData(x, TTList_TTMaxGSV), ZeroValue) & ", " & _
                """" & arrData(x, TTList_TTCondComment) & """, " & _
                SetEmptyValue(arrData(x, TTList_StandardTerm), ZeroValue) & ", " & _
                SetEmptyValue(arrData(x, TTList_AddnlTerm), ZeroValue) & ", " & _
                SetEmptyValue(arrData(x, TTList_BannerTerm), ZeroValue) & ")"
98        cn.Execute qry
99    Next x


      ' Refresh Trading Terms Const
      ' Delete all records with the Ref#
100   qry = vbNullString
101   qry = "DELETE * " & _
            "FROM " & OP_TRADING_TERMS_CONST_TBL & " " & _
            "WHERE RefNumber = '" & strRefNumber & "';"
102   cn.Execute qry

103   qry = "INSERT INTO " & OP_TRADING_TERMS_CONST_TBL & " " & _
            "(RefNumber, AllProd, AllProd_PctGSVlessQA3, AllProd_DollarperLtr, AllProd_FreqOfPayments, AllNonContrdProd, AllNonContrdProd_PctGSVlessQA3, AllNonContrdProd_DollarperLtr, BannerTerms, BannerTerms_PctGSV, BannerTerms_PctGSVlessQA3) " & _
            "VALUES ( " & _
            """" & strRefNumber & """, " & _
            IIf(frm.optContractAndNonContract.Value = True, True, False) & ", " & _
            SetEmptyValue(frm.txtAllProd_PctGSV, NullForDB) & ", " & _
            SetEmptyValue(frm.txtAllProd_DollarPerLitre, NullForDB) & ", " & _
            """" & frm.cboAllProd_FreqOfPayments & """, " & _
            IIf(frm.chkNonContract.Value = True, True, False) & ", " & _
            SetEmptyValue(frm.txtAllNonContract_PctGSV, NullForDB) & ", " & _
            SetEmptyValue(frm.txtAllNonContract_DollarPerLitre, NullForDB) & ", " & _
            IIf(frm.chkBannerTerms.Value = True, True, False) & ", " & _
            SetEmptyValue(frm.txtTTBannerGSV, NullForDB) & ", " & _
            SetEmptyValue(frm.txtTTBannerGSVlessQA3, NullForDB) & ")"
104   cn.Execute qry


      ' Refresh COOP and A&P
      ' Delete all records with the Ref#
105   qry = "DELETE * " & _
            "FROM " & OP_COOP_ANP_TBL & " " & _
            "WHERE RefNumber = '" & strRefNumber & "';"
106   cn.Execute qry

107   qry = "INSERT INTO " & OP_COOP_ANP_TBL & " " & _
            "(RefNumber, CashPaymentCoop, BonusStockCoop, PromoFundCoop, StaffIncentivesCoop, PRAHospitalityCoop, CashPaymentAnP, " & _
             "BonusStockAnP, PromoFundAnP, StaffIncentivesAnP, PRAHospitalityAnP, CashPaymentComments, BonusStockComments, PromoFundComments, " & _
             "StaffIncentivesComments, PRAHospitalityComments, ReciprocalSpend, ReciprocalSpendComments) " & _
            "VALUES ( " & _
            """" & strRefNumber & """, " & _
            SetEmptyValue(frm.txtCoopCashPay, ZeroValue) & ", " & _
            SetEmptyValue(frm.txtCoopBonusStock, ZeroValue) & ", " & _
            SetEmptyValue(frm.txtCoopPromoFund, ZeroValue) & ", " & _
            SetEmptyValue(frm.txtCoopStaffIncentives, ZeroValue) & ", " & _
            SetEmptyValue(frm.txtCoopPRAHospitality, ZeroValue) & ", " & _
            SetEmptyValue(frm.txtAnPCashPay, ZeroValue) & ", " & _
            SetEmptyValue(frm.txtAnPBonusStock, ZeroValue) & ", " & _
            SetEmptyValue(frm.txtAnPPromoFund, ZeroValue) & ", " & _
            SetEmptyValue(frm.txtAnPStaffIncentives, ZeroValue) & ", " & _
            SetEmptyValue(frm.txtAnPPRAHospitality, ZeroValue) & ", " & _
            """" & SetEmptyValue(frm.txtCommentsCashPay, NullStrings) & """, " & _
            """" & SetEmptyValue(frm.txtCommentsBonusStock, NullStrings) & """, " & _
            """" & SetEmptyValue(frm.txtCommentsPromoFund, NullStrings) & """, " & _
            """" & SetEmptyValue(frm.txtCommentsStaffIncentives, NullStrings) & """, " & _
            """" & SetEmptyValue(frm.txtCommentsPRAHospitality, NullStrings) & """, " & _
            SetEmptyValue(frm.txtReciprocalSpend, ZeroValue) & ", " & _
            """" & SetEmptyValue(frm.txtReciprocalSpendComments, NullStrings) & """)"
108   cn.Execute qry


      ' Update last sync local date
109   qry = vbNullString
110   qry = qry & "UPDATE " & SYNC_DATE_TBL & " " & _
                  "SET LastSyncDate = #" & ConvertLocalToGMT(Now) & "# " & _
                  "WHERE ID = '" & GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "ID", "WinLoginName", UCase(Environ("UserName")), """") & "'"
111   cn.Execute qry

Proc_Exit:
112   PopCallStack
113   Exit Function

Err_Handler:
114   GlobalErrHandler
115   Resume Proc_Exit
          
End Function

Public Sub PopulateOPDetails(strRefNum As String, frm As frmMain)
      Dim qry As String
      Dim rs As ADODB.Recordset
      Dim ws As Worksheet
      Dim rng As Range
      Dim arrData As Variant
      Dim x As Integer, y As Integer

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mOPInputs|PopulateOPDetails"

3     Set rs = New ADODB.Recordset

      ' Get data from Main table
4     qry = "SELECT * FROM " & OP_MAIN_TBL & " WHERE RefNumber = '" & strRefNum & "'"
5     rs.Open qry, cn

6     If Not rs.EOF Then
          ' Ref#
7         frm.txtRefNumber = strRefNum
          
          ' Status ID
8         frm.lblStat.Caption = rs.Fields("StatusID").Value
          
          ' PRA Contact Information
9         frm.cboCreator.Value = rs.Fields("CreatorID").Value
          
          ' PROS Segmentation
10        frm.cboPROS = rs.Fields("PROS").Value
          
          ' Customer Contact Information
11        frm.txtAddress = rs.Fields("CustomerAddress").Value
12        frm.txtCustomerName = rs.Fields("CustomerName").Value
13        frm.txtCustomerPhone = rs.Fields("CustomerPhone").Value
14        frm.txtCustomerEmail = rs.Fields("CustomerEmail").Value
          
          ' Contract Level
15        frm.cboContractLevel.Value = rs.Fields("ContractLevel").Value
16        Call LBXSelectItems(frm.lstContractLevelCode, Split(SetEmptyValue(rs.Fields("ContractLevelCode").Value, NullStrings), "|"), 1)
17        frm.txtOutletOrGroupName = rs.Fields("OutletOrGroupName").Value
                
          ' Contract Date
18        frm.txtFromDate = Format(rs.Fields("FromDate").Value, "dd-mmm-yyyy")
19        frm.txtToDate = Format(rs.Fields("ToDate").Value, "dd-mmm-yyyy")
20        frm.txtFromExtnDate = Format(rs.Fields("FromDate_Extention").Value, "dd-mmm-yyyy")
21        frm.txtToExtnDate = Format(rs.Fields("ToDate_Extention").Value, "dd-mmm-yyyy")

          ' Route to Market
22        frm.cboRouteToMarket.Value = rs.Fields("RouteToMarket").Value
          '---ws_code should be in col 1
23        Call LBXSelectItems(frm.lstWholesaler, Split(SetEmptyValue(rs.Fields("Wholesaler").Value, NullStrings), "|"))
          
          ' Contract Type
24        frm.chkSpirits.Value = IIf(InStr(1, rs.Fields("ContractType").Value, "Spirits") <> 0, True, False)
25        frm.chkChampagne.Value = IIf(InStr(1, rs.Fields("ContractType").Value, "Champagne") <> 0, True, False)
26        frm.chkWine.Value = IIf(InStr(1, rs.Fields("ContractType").Value, "Wine") <> 0, True, False)
          
          ' Contract Form
27        frm.cboContractForm.Value = rs.Fields("ContractForm").Value
          
          ' Variable Price
28        frm.cboVarPrice.Value = rs.Fields("VarPrice").Value
          
          ' Comments
29        frm.txtComments = rs.Fields("Comments").Value

30    End If
31    Call CloseRecordset(rs)


      ' Product list
32    qry = "SELECT DISTINCT T1.ProductType, T2.BRAND_NAME, T1.BrandCode, T2.SUB_BRAND_NAME, T1.SubBrandCode, " & _
              "T2.PRODUCT_DESCRIPTION, T1.ProductCode, T1.BottleSize, T1.UnitsPerCase, " & _
              "T1.ContractedCases, T1.ContractedVolume, ROUND(T1.ContractedVolume,0), T1.ContractedGSV, ROUND(T1.ContractedGSV,0), " & _
              "T1.Family " & _
            "FROM " & OP_PROD_DETAILS_TBL & " AS T1 LEFT JOIN " & PRODUCT_MAP_TBL & " AS T2 " & _
              "ON (T1.UnitsPerCase = T2.UNITS_PER_CASE) " & _
             "AND (T1.BottleSize = T2.BOTTLE_SIZE) " & _
             "AND (T1.ProductCode = T2.PRODUCT_CODE) " & _
             "AND (T1.SubBrandCode = T2.SUB_BRAND_CODE) " & _
             "AND (T1.BrandCode = T2.BRAND_CODE) " & _
            "WHERE T1.RefNumber = '" & strRefNum & "' " & _
            "ORDER BY T2.PRODUCT_DESCRIPTION"
33    rs.Open qry, cn
34    If Not rs.EOF Then frm.lstProducts.List = ConvertSQLArrToListArr(rs.GetRows(10000))
35    Call CloseRecordset(rs)

      ' Trading Terms list
36    qry = "SELECT DISTINCT T3.ProductType, T2.BRAND_NAME, T2.PRODUCT_DESCRIPTION, T1.ProductCode, ROUND(T3.ContractedVolume,0), ROUND(T3.ContractedGSV,0), IIF(T1.DollarPerLiter=0,'',T1.DollarPerLiter), IIF(T1.PctOfGSV=0,'',T1.PctOfGSV), " & _
              "T1.FreqOfPayments, IIF(T1.AddnlDollarPerLiter=0,'',T1.AddnlDollarPerLiter), IIF(T1.AddnlPctOfGSV=0,'',T1.AddnlPctOfGSV), " & _
              "T1.CondTermComments, T1.StandardTerms, T1.AdditionalTerms, T1.BannerTerms " & _
            "FROM (" & OP_TRADING_TERMS_TBL & " AS T1 LEFT JOIN " & PRODUCT_MAP_TBL & " AS T2 ON T1.ProductCode = T2.PRODUCT_CODE) " & _
            "INNER JOIN " & OP_PROD_DETAILS_TBL & " AS T3 ON T1.RefNumber = T3.RefNumber AND T1.ProductCode = T3.ProductCode " & _
            "WHERE T1.RefNumber = '" & strRefNum & "' " & _
            "ORDER BY T2.PRODUCT_DESCRIPTION"
37    rs.Open qry, cn
38    If Not rs.EOF Then frm.lstTrdTerms.List = ConvertSQLArrToListArr(rs.GetRows(10000))
39    Call CloseRecordset(rs)


      ' QA3 list
40    qry = "SELECT DISTINCT T1.ProductType, T2.BRAND_NAME, T2.PRODUCT_DESCRIPTION, T1.ProductCode, ROUND(T1.ContractedVolume,0), ROUND(T1.ContractedGSV,0), " & _
              "IIF(T1.DirectPrice=0,'',T1.DirectPrice), IIF(T1.DirectPrice=0,'',ROUND(T1.DirectPrice,2)), " & _
              "IIF(T1.WholesalePrice=0,'',T1.WholesalePrice), IIF(T1.WholesalePrice=0,'',ROUND(T1.WholesalePrice,2)), " & _
              "IIF(T1.QA3PerCaseUser=0,'',T1.QA3PerCaseUser), IIF(T1.QA3PerCaseUser=0,'',ROUND(T1.QA3PerCaseUser,2)), " & _
              "IIF(T1.NIPOrLUCAuto=0,'',T1.NIPOrLUCAuto), IIF(T1.NIPOrLUCAuto=0,'',ROUND(T1.NIPOrLUCAuto,2)), " & _
              "IIF(T1.NIPOrLUCUser=0,'',T1.NIPOrLUCUser), IIF(T1.NIPOrLUCUser=0,'',ROUND(T1.NIPOrLUCUser,2)), " & _
              "IIF(T1.QA3PerCaseAuto=0,'',T1.QA3PerCaseAuto), IIF(T1.QA3PerCaseAuto=0,'',ROUND(T1.QA3PerCaseAuto,2)), " & _
              "IIF(T1.QA3=0,'',T1.QA3), IIF(T1.QA3=0,'',ROUND(T1.QA3,2)), " & _
              "IIF(T1.KWI=0,'',T1.KWI), IIF(T1.KWI=0,'',ROUND(T1.KWI,2)), " & _
              "IIF(T1.COP=0,'',T1.COP), IIF(T1.COP=0,'',ROUND(T1.COP,2)), " & _
              "T1.Family " & _
            "FROM " & OP_PROD_DETAILS_TBL & " AS T1 LEFT JOIN " & PRODUCT_MAP_TBL & " AS T2 ON (T1.UnitsPerCase = T2.UNITS_PER_CASE) " & _
             "AND (T1.BottleSize = T2.BOTTLE_SIZE) AND (T1.ProductCode = T2.PRODUCT_CODE) AND (T1.SubBrandCode = T2.SUB_BRAND_CODE) AND (T1.BrandCode = T2.BRAND_CODE) " & _
            "WHERE T1.RefNumber = '" & strRefNum & "' " & _
            "ORDER BY T2.PRODUCT_DESCRIPTION"
41    rs.Open qry, cn
42    If Not rs.EOF Then frm.lstQA3.List = ConvertSQLArrToListArr(rs.GetRows(10000))
43    Call CloseRecordset(rs)


      ' Trading Terms Const
44    qry = "SELECT DISTINCT AllProd, AllProd_PctGSVlessQA3, AllProd_DollarperLtr, AllProd_FreqOfPayments, AllNonContrdProd, AllNonContrdProd_PctGSVlessQA3, AllNonContrdProd_DollarperLtr, BannerTerms, BannerTerms_PctGSV, BannerTerms_PctGSVlessQA3 " & _
            "FROM " & OP_TRADING_TERMS_CONST_TBL & " " & _
            "WHERE RefNumber = '" & strRefNum & "'"
45    rs.Open qry, cn

46    If Not rs.EOF Then
47        frm.optContractAndNonContract = rs.Fields("AllProd").Value
48        frm.txtAllProd_PctGSV = SetEmptyValue(rs.Fields("AllProd_PctGSVlessQA3").Value, NullStrings)
49        frm.txtAllProd_DollarPerLitre = SetEmptyValue(rs.Fields("AllProd_DollarperLtr").Value, NullStrings)
50        frm.cboAllProd_FreqOfPayments = rs.Fields("AllProd_FreqOfPayments").Value
51        frm.chkNonContract = rs.Fields("AllNonContrdProd").Value
52        frm.txtAllNonContract_PctGSV = SetEmptyValue(rs.Fields("AllNonContrdProd_PctGSVlessQA3").Value, NullStrings)
53        frm.txtAllNonContract_DollarPerLitre = SetEmptyValue(rs.Fields("AllNonContrdProd_DollarperLtr").Value, NullStrings)
54        frm.chkBannerTerms = rs.Fields("BannerTerms").Value
55        frm.txtTTBannerGSV = SetEmptyValue(rs.Fields("BannerTerms_PctGSV").Value, NullStrings)
56        frm.txtTTBannerGSVlessQA3 = SetEmptyValue(rs.Fields("BannerTerms_PctGSVlessQA3").Value, NullStrings)
57    End If
58    Call CloseRecordset(rs)


      ' Coop and A&P
59    qry = "SELECT RefNumber, IIF(CashPaymentCoop=0,'',CashPaymentCoop) AS CoopCashPayment, IIF(BonusStockCoop=0,'',BonusStockCoop) AS CoopBonusStock, IIF(PromoFundCoop=0,'',PromoFundCoop) AS CoopPromoFund, IIF(StaffIncentivesCoop=0,'',StaffIncentivesCoop) AS CoopStaffIncentives, IIF(PRAHospitalityCoop=0,'',PRAHospitalityCoop) AS CoopPRAHospitality, " & _
              "IIF(CashPaymentAnP=0,'',CashPaymentAnP) AS AnPCashPayment, IIF(BonusStockAnP=0,'',BonusStockAnP) AS AnPBonusStock, IIF(PromoFundAnP=0,'',PromoFundAnP) AS AnPPromoFund, IIF(StaffIncentivesAnP=0,'',StaffIncentivesAnP) AS AnPStaffIncentives, IIF(PRAHospitalityAnP=0,'',PRAHospitalityAnP) AS AnPPRAHospitality, " & _
              "CashPaymentComments, BonusStockComments, PromoFundComments, StaffIncentivesComments, PRAHospitalityComments, IIF(ReciprocalSpend=0,'',ReciprocalSpend) AS Reciprocal_Spend, ReciprocalSpendComments " & _
            "FROM " & OP_COOP_ANP_TBL & " " & _
            "WHERE RefNumber = '" & strRefNum & "'"
60    rs.Open qry, cn

61    If Not rs.EOF Then
62        frm.txtCoopCashPay = rs.Fields("CoopCashPayment").Value
63        frm.txtCoopBonusStock = rs.Fields("CoopBonusStock").Value
64        frm.txtCoopPromoFund = rs.Fields("CoopPromoFund").Value
65        frm.txtCoopStaffIncentives = rs.Fields("CoopStaffIncentives").Value
66        frm.txtCoopPRAHospitality = rs.Fields("CoopPRAHospitality").Value
67        frm.txtAnPCashPay = rs.Fields("AnPCashPayment").Value
68        frm.txtAnPBonusStock = rs.Fields("AnPBonusStock").Value
69        frm.txtAnPPromoFund = rs.Fields("AnPPromoFund").Value
70        frm.txtAnPStaffIncentives = rs.Fields("AnPStaffIncentives").Value
71        frm.txtAnPPRAHospitality = rs.Fields("AnPPRAHospitality").Value
72        frm.txtCommentsCashPay = rs.Fields("CashPaymentComments").Value
73        frm.txtCommentsBonusStock = rs.Fields("BonusStockComments").Value
74        frm.txtCommentsPromoFund = rs.Fields("PromoFundComments").Value
75        frm.txtCommentsStaffIncentives = rs.Fields("StaffIncentivesComments").Value
76        frm.txtCommentsPRAHospitality = rs.Fields("PRAHospitalityComments").Value
77        frm.txtReciprocalSpend = rs.Fields("Reciprocal_Spend").Value
78        frm.txtReciprocalSpendComments = rs.Fields("ReciprocalSpendComments").Value
79    End If
80    Call CloseRecordset(rs, True)

Proc_Exit:
81    PopCallStack
82    Exit Sub

Err_Handler:
83    GlobalErrHandler
84    Resume Proc_Exit
End Sub
