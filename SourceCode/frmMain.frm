VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "On Premise Tool"
   ClientHeight    =   10380
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18120
   OleObjectBlob   =   "frmMain.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Dim blnFilterUpdating As Boolean

Private Sub cboContractForm_Change()

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cboContractForm_Change"

3     Select Case cboContractForm
          Case "Short Form"
4             Call EnableTextBox(txtAnPCashPay, False)
5             Call EnableTextBox(txtAnPBonusStock, False)
6             Call EnableTextBox(txtAnPPromoFund, False)
7             Call EnableTextBox(txtAnPStaffIncentives, False)
8             Call EnableTextBox(txtAnPPRAHospitality, False)
9         Case "Long Form"
10            Call EnableTextBox(txtAnPCashPay)
11            Call EnableTextBox(txtAnPBonusStock)
12            Call EnableTextBox(txtAnPPromoFund)
13            Call EnableTextBox(txtAnPStaffIncentives)
14            Call EnableTextBox(txtAnPPRAHospitality)
15    End Select

Proc_Exit:
16    PopCallStack
17    Exit Sub

Err_Handler:
18    GlobalErrHandler
19    Resume Proc_Exit
End Sub


Private Sub cmdMonthlyReport_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdMonthlyReport_Click"

3     frmMonthlyReport.Show

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub CommandButton1_Click()
1         Call CopyTableFromDB("C:\Users\Public\Documents\Projects\AUS OP\Mappings - Copy.accdb", "Table1", "prau$", _
                               "C:\Users\Public\Documents\Projects\AUS OP\Mappings.accdb", "Table1", "prau$")
End Sub

Private Sub cmdDataDump_Click()
      Dim wb As Workbook

      ' View only if status is submitted for approval
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdDataDump_Click"

3     If lstRecord.ListCount <> 0 Then
4         Set wb = CreateViewWorkbook(Array(DATA_DUMP_SHEET))

5         Call CreateDataDumpReport(wb)

          ' Delete templates
6         Call DeleteViewTemplates(wb, Array(DATA_DUMP_SHEET))
          
7         Set wb = Nothing
8     End If

Proc_Exit:
9     PopCallStack
10    Exit Sub

Err_Handler:
11    GlobalErrHandler
12    Resume Proc_Exit

End Sub

Private Sub mpgeOP_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|mpgeOP_Change"

3     If mpgeOP.SelectedItem.Name = "pgePEMPreview" Then
4         Application.Wait Now + TimeValue("0:00:01")
      '    Application.Cursor = xlWait
5         If lstProducts.ListCount > 0 Then
6             Call PEM_Preview(wkb, Me)
7         End If
      '    Application.Cursor = xlDefault
8     End If

Proc_Exit:
9     PopCallStack
10    Exit Sub

Err_Handler:
11    GlobalErrHandler
12    Resume Proc_Exit
End Sub

Private Sub UserForm_Initialize()

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|UserForm_Initialize"

3     Call InitCommonControls

      ' Set current workbook
4     Set wkb = ThisWorkbook

      ' Create Min and Max button in form
5     Call FormatUserForm(Me.Caption)

      ' Setup database connection
6     Call SetDBConnection(cn)

      ' Get user login ID
7     g_sLoginID = UCase(Environ("UserName"))

      ' Get user access priviledges
8     g_iAccessType = CInt(GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "AccessTypeID", "WinLoginName", g_sLoginID, """"))
      ' Hide Sync button if user is admin
9     Select Case g_iAccessType
          Case enumUserPermission.OrdinaryUser
10            cmdImportMDMData.Visible = False
11            cmdSyncDB.Visible = True
12            lblLoginInfo.Caption = lblLoginInfo.Caption & " Ordinary User"
13        Case enumUserPermission.Admin
14            cmdImportMDMData.Visible = True
15            cmdSyncDB.Visible = False
16            lblLoginInfo.Caption = lblLoginInfo.Caption & " Administrator"
17        Case enumUserPermission.Manager
18            cmdImportMDMData.Visible = False
19            cmdSyncDB.Visible = False
20            lblLoginInfo.Caption = lblLoginInfo.Caption & " Manager"
21    End Select

      ' Get calculation Constants
22    g_dblNIP_Const = GetItemFromMappingTbl(SETTINGS_TBL, "Settings_Value", "Settings_Name", "NIP_Constant", """")
23    g_dblWET = GetItemFromMappingTbl(SETTINGS_TBL, "Settings_Value", "Settings_Name", "WET", """")
24    g_dblALM_Admin = GetItemFromMappingTbl(SETTINGS_TBL, "Settings_Value", "Settings_Name", "ALM_Admin", """")
25    g_dblALM_Freight = GetItemFromMappingTbl(SETTINGS_TBL, "Settings_Value", "Settings_Name", "ALM_Freight", """")

      ' Populate list
26    Call PopulateFrontPageList(lstRecord)

      ' OP Contract Parameters --------------------------------------------------------------Start
      ' Hide details
27    Call HideDetails

      ' PRA Contact Information
28    cboCreator.List = GetArrayList("SELECT DISTINCT [Name], ID FROM " & PRA_EMPLOYEE_TBL & ";", True)
29    cboCreator.Text = GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "[Name]", "WinLoginName", g_sLoginID, "'")

      ' Contract Level
30    With cboContractLevel
31        .AddItem "OP Banner"
32        .AddItem "OP Banner Region"
33        .AddItem "OP Outlet Level"
34    End With

      ' Route to Market
35    With cboRouteToMarket
36        .AddItem "Direct"
37        .AddItem "Indirect"
38    End With

      ' Contract Form
39    With cboContractForm
40        .AddItem "Short Form"
41        .AddItem "Long Form"
42    End With
      ' OP Contract Parameters --------------------------------------------------------------End

      ' Fill record list
43    Call PopulateProductSelection

      ' Populate Frequency Of Payments
44    cboAllProd_FreqOfPayments.List = Split(GetItemFromMappingTbl(SETTINGS_TBL, "Settings_Value", "Settings_Name", "Freq_of_Payments", """"), "|")
45    cboTTFreqOfPayments.List = Split(GetItemFromMappingTbl(SETTINGS_TBL, "Settings_Value", "Settings_Name", "Freq_of_Payments", """"), "|")

      ' PROS Segmentation
46    cboPROS.List = Split(GetItemFromMappingTbl(SETTINGS_TBL, "Settings_Value", "Settings_Name", "PROS_Segmentation", """"), "|")

      ' Clear temp sheets
47    Call ClearTempSheet

Proc_Exit:
48    PopCallStack
49    Exit Sub

Err_Handler:
50    GlobalErrHandler
51    Resume Proc_Exit
End Sub

Private Sub ClearTempSheet()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|ClearTempSheet"

3     wkb.Worksheets(PEM_PREVIW_TEMP_SHEET).Cells.ClearContents

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub cmdImportMDMData_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdImportMDMData_Click"

3     Call TransferMDMData

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

' Contract Level------------------------------------------Start
Private Sub cboContractLevel_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cboContractLevel_Change"

      ' Clear Outlet/Group Name
3     txtOutletOrGroupName = vbNullString

      ' Populate code list
4     Select Case cboContractLevel.Value
          Case "OP Banner"
5             lstContractLevelCode.List = GetArrayList("SELECT DISTINCT Banner, BannerCode FROM " & CUSTOMER_MAP_TBL & ";", True)
6             lstContractLevelCode.MultiSelect = fmMultiSelectSingle
7             lstContractLevelCode.ColumnCount = 2
8             lstContractLevelCode.ColumnWidths = "246;56"
9             lblCode.Caption = "  Banner Code"
10            lblSiteCode.Visible = False
11        Case "OP Banner Region"
12            lstContractLevelCode.List = GetArrayList("SELECT DISTINCT BannerRegion, BannerRegionCode FROM " & CUSTOMER_MAP_TBL & ";", True)
13            lstContractLevelCode.MultiSelect = fmMultiSelectSingle
14            lstContractLevelCode.ColumnCount = 2
15            lstContractLevelCode.ColumnWidths = "246;56"
16            lblCode.Caption = "  BRC Code"
17            lblSiteCode.Visible = False
18        Case "OP Outlet Level"
19            lstContractLevelCode.List = GetArrayList(AddDummyOutlet(CUSTOMER_MAP_TBL) & "SELECT DISTINCT OutletName, ExternalID, MatchCode FROM " & CUSTOMER_MAP_TBL & ";", True)
20            lstContractLevelCode.MultiSelect = fmMultiSelectMulti
21            lstContractLevelCode.ColumnCount = 3
22            lstContractLevelCode.ColumnWidths = "180;66;56"
23            lblCode.Caption = "  Match Code"
24            lblSiteCode.Visible = True
25    End Select

Proc_Exit:
26    PopCallStack
27    Exit Sub

Err_Handler:
28    GlobalErrHandler
29    Resume Proc_Exit
End Sub
' Contract Level------------------------------------------End

' Route to Market-----------------------------------------Start
Private Sub cboRouteToMarket_Enter()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cboRouteToMarket_Enter"

      ' Items in the QA3 Only tab will be cleared if changed
3     If lstProducts.ListCount > 0 Then
4         MsgBox "The items in QA3 Only tab will be cleared if you change the Route To Market.", vbInformation
5     End If

Proc_Exit:
6     PopCallStack
7     Exit Sub

Err_Handler:
8     GlobalErrHandler
9     Resume Proc_Exit
End Sub

Private Sub cboRouteToMarket_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cboRouteToMarket_Change"

      ' Delete listbox items
3     lstProducts.Clear
4     lstTrdTerms.Clear
5     lstQA3.Clear
6     Call ClearTempSheet

      ' Clear items
7     lstWholesaler.Clear

8     Select Case cboRouteToMarket.Value
          Case "Indirect"
9             lstWholesaler.List = GetArrayList("SELECT Code, ItemName FROM " & WHOLESALER_MAP_TBL & ";", True)

10            With cboVarPrice
11                .Clear
12                .AddItem "900+"
13                .Text = "900+"
14                .Enabled = False
15            End With

16            txtDirectPrice = vbNullString
17            Call EnableTextBox(txtWholesalerPrice)
18            txtWholesalerPrice = vbNullString

19        Case "Direct"
20            With cboVarPrice
21                .Clear
22                .AddItem "0-9"
23                .AddItem "10-99"
24                .AddItem "100-499"
25                .AddItem "500-899"
26                .AddItem "900+"
27                .Enabled = True
28            End With

29            txtDirectPrice = vbNullString
30            Call EnableTextBox(txtWholesalerPrice, False)
31            txtWholesalerPrice = vbNullString

32        Case ""
              ' nothing
33    End Select

Proc_Exit:
34    PopCallStack
35    Exit Sub

Err_Handler:
36    GlobalErrHandler
37    Resume Proc_Exit
End Sub
' Route to Market-----------------------------------------End

' Record filtering----------------------------------------Start
Private Sub imgFilterRefNum_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|imgFilterRefNum_Click"

3     Call CreateFilterPopUpMenu(lstRecord, 0)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub imgFilterCreator_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|imgFilterCreator_Click"

3     Call CreateFilterPopUpMenu(lstRecord, 1)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub imgFilterAuthoriser_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|imgFilterAuthoriser_Click"

3     Call CreateFilterPopUpMenu(lstRecord, 2)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub imgFilterContractLevel_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|imgFilterContractLevel_Click"

3     Call CreateFilterPopUpMenu(lstRecord, 3)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub imgFilterOutletGroupName_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|imgFilterOutletGroupName_Click"

3     Call CreateFilterPopUpMenu(lstRecord, 4)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub imgFilterStartDate_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|imgFilterStartDate_Click"

3     Call CreateFilterPopUpMenu(lstRecord, 5)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub imgFilterEndDate_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|imgFilterEndDate_Click"

3     Call CreateFilterPopUpMenu(lstRecord, 6)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub imgFilterSubmitDate_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|imgFilterSubmitDate_Click"

3     Call CreateFilterPopUpMenu(lstRecord, 7)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub imgFilterStatus_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|imgFilterStatus_Click"

3     Call CreateFilterPopUpMenu(lstRecord, 8)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub
' Record filtering----------------------------------------End

Private Sub cmdSyncDB_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdSyncDB_Click"

3     Call SyncDatabase

      ' Repopulate list to reflect any changes
4     Call PopulateFrontPageList(lstRecord)

Proc_Exit:
5     PopCallStack
6     Exit Sub

Err_Handler:
7     GlobalErrHandler
8     Resume Proc_Exit
End Sub

Private Sub cmdCreateNew_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdCreateNew_Click"

      ' Clear contents
3     Call ClearDetails
4     Call RefreshSelection

      ' Reset Var Price List
5     With cboVarPrice
6         .Clear
7         .Enabled = False
8     End With

      ' Hide controls
9     cmdApprove.Visible = False
10    cmdReplicate.Visible = False
11    cmdDelete.Visible = False

12    Call UserPermissionSettings

      ' Generate Ref#
13    txtRefNumber = GenerateRefNum

      ' Clear temp sheets
14    Call ClearTempSheet

      ' Show details
15    Call ShowDetails

Proc_Exit:
16    PopCallStack
17    Exit Sub

Err_Handler:
18    GlobalErrHandler
19    Resume Proc_Exit
End Sub

Private Sub UserPermissionSettings()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|UserPermissionSettings"

3     If g_iAccessType = enumUserPermission.Admin Or g_iAccessType = enumUserPermission.Manager Then
          ' Enable editing of creator combobox
4         cboCreator.Locked = False
          ' Show dropdown button of creator combobox
5         cboCreator.ShowDropButtonWhen = fmShowDropButtonWhenAlways
6     End If

Proc_Exit:
7     PopCallStack
8     Exit Sub

Err_Handler:
9     GlobalErrHandler
10    Resume Proc_Exit
End Sub

Private Sub cmdViewRecord_Click()
      Dim strStatus As String

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdViewRecord_Click"

3     If LBXSelectCount(Me.lstRecord) <> 0 Then
          ' Set focus to current workbook
          'wkb.Activate
          
          ' Clear contents
4         Call ClearDetails
5         Call RefreshSelection
          
          ' Show details
6         Call ShowDetails
          
          ' Get Status
7         strStatus = lstRecord.List(lstRecord.ListIndex, 8)
          
          ' Show command buttons
8         cmdReplicate.Visible = True
9         cmdDelete.Visible = True
                 
10        If (g_iAccessType = enumUserPermission.Admin Or g_iAccessType = enumUserPermission.Manager) And strStatus = "For Approval" Then
11            cmdApprove.Visible = True
12        Else
13            cmdApprove.Visible = False
14        End If
          
          ' Populate GUI
15        Call PopulateOPDetails(lstRecord.List(lstRecord.ListIndex, 0), Me)
16    End If

Proc_Exit:
17    PopCallStack
18    Exit Sub

Err_Handler:
19    GlobalErrHandler
20    Resume Proc_Exit
End Sub

Private Sub RefreshSelection()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|RefreshSelection"

      ' PRA Contact Information
3     cboCreator.List = GetArrayList("SELECT DISTINCT [Name], ID FROM " & PRA_EMPLOYEE_TBL & ";", True)
4     cboCreator.Text = GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "[Name]", "WinLoginName", g_sLoginID, "'")

Proc_Exit:
5     PopCallStack
6     Exit Sub

Err_Handler:
7     GlobalErrHandler
8     Resume Proc_Exit
End Sub

Private Sub ShowDetails()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|ShowCoopDetails"

      ' Show the 1st page
3     mpgeOP.Value = 0

      ' Disable controls
4     fraOP_List.Enabled = False

      ' Display COOP details
5     fraDetailsFront.Visible = True
6     fraDetailsBack.Visible = False

Proc_Exit:
7     PopCallStack
8     Exit Sub

Err_Handler:
9     GlobalErrHandler
10    Resume Proc_Exit
End Sub

Private Sub HideDetails()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|HideDetails"

      ' Disable editing of creator combobox
3     cboCreator.Locked = True
      ' Hide dropdown button of creator combobox
4     cboCreator.ShowDropButtonWhen = fmShowDropButtonWhenNever

      ' Enable controls
5     fraOP_List.Enabled = True

      ' Hide details
6     fraDetailsFront.Visible = False
7     fraDetailsBack.Visible = True

      ' Populate list
8     Call PopulateFrontPageList(lstRecord)

Proc_Exit:
9     PopCallStack
10    Exit Sub

Err_Handler:
11    GlobalErrHandler
12    Resume Proc_Exit
End Sub

Private Sub ClearDetails()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|ClearDetails"

      ' OP Contract Parameters
3     txtRefNumber = vbNullString
4     cboCreator.Value = vbNullString
5     cboManager.Text = vbNullString
6     txtOutletOrGroupName = vbNullString
7     txtAddress = vbNullString
8     txtCustomerName = vbNullString
9     txtCustomerPhone = vbNullString
10    txtCustomerEmail = vbNullString
11    cboContractLevel.Value = vbNullString
12    lstContractLevelCode.Clear
13    lstContractLevelCode.ColumnCount = 2
14    lstContractLevelCode.ColumnWidths = "246;56"
15    lblCode.Caption = "  Code"
16    lblSiteCode.Visible = False
17    cboPROS = vbNullString
18    txtFromDate = vbNullString
19    txtToDate = vbNullString
20    txtFromExtnDate = vbNullString
21    txtToExtnDate = vbNullString
22    cboRouteToMarket = vbNullString
23    lstWholesaler.Clear
24    chkSpirits.Value = False
25    chkChampagne = False
26    chkWine = False
27    cboContractForm = vbNullString
      'cboVarPrice.Clear
28    txtComments = vbNullString

29    Call ClearTempSheet

      ' Products
30    Call cmdClearInputsProd_Click
31    lstProducts.Clear

      ' Trading Terms
32    Call cmdClearInputsTradingTerms_Click
33    lstTrdTerms.Clear
34    optEnterManually.Value = True
35    chkNonContract.Value = False
36    txtAllProd_PctGSV = vbNullString
37    txtAllProd_DollarPerLitre = vbNullString
38    cboAllProd_FreqOfPayments = vbNullString
39    txtAllNonContract_PctGSV = vbNullString
40    txtAllNonContract_DollarPerLitre = vbNullString
41    txtAllCondMaxGSV = vbNullString
42    txtAllCondMaxLitre = vbNullString
43    txtAllCondComment = vbNullString
44    chkBannerTerms.Value = False
45    txtTTBannerGSV = vbNullString
46    txtTTBannerGSVlessQA3 = vbNullString

      ' QA3
47    Call cmdClearInputsQA3_Click
48    lstQA3.Clear

      ' COOP and A&P
49    Call cmdClearCOOPAndAnP_Click

      '' PEM preview
      'txtPEM_Total_Vol = 0
      'txtPEM_Total_GSV = 0
      'txtPEM_Total_KWI = 0
      'txtPEM_Total_QA3 = 0
      'txtPEM_Total_Terms = 0
      'txtPEM_Total_COOP = 0
      'txtPEM_Total_Other = 0
      'txtPEM_Total_NSV = 0
      'txtPEM_Total_COGS = 0
      'txtPEM_Total_CM = 0
      'txtPEM_Total_AnD_GSV = 0
      'txtPEM_Total_LUC = 0
      'txtPEM_Total_NIP = 0
      'txtPEM_Total_CM_NSV = 0
      'lstPEMPreview.Clear

Proc_Exit:
50    PopCallStack
51    Exit Sub

Err_Handler:
52    GlobalErrHandler
53    Resume Proc_Exit
End Sub

Private Sub UserForm_Terminate()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|UserForm_Terminate"

3     Call CloseDBConnection(cn)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

' COOP details-----------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------

Private Sub cboCreator_Change()
      Dim intManagerID As Integer

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cboCoopCreator_Change"

3     If Len(cboCreator.Value) <> 0 Then
          ' Change manager
4         intManagerID = GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "ManagerID", "ID", cboCreator.Value, """")
5         cboManager.Value = GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "Name", "ID", GetItemFromMappingTbl(PRA_MANAGER_TBL, "[Name]", "ID", CStr(intManagerID)), """")
          
          ' Change COOP Ref# prefix
6         If Len(txtRefNumber) <> 0 Then
7             txtRefNumber.Value = cboCreator.Value & Mid(txtRefNumber.Value, InStr(1, txtRefNumber.Value, "-"))
8         End If
9     End If

Proc_Exit:
10    PopCallStack
11    Exit Sub

Err_Handler:
12    GlobalErrHandler
13    Resume Proc_Exit
End Sub

' Dates---------------------------------------------------Start
Private Sub txtFromDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtCoopFromDate_MouseDown"

3     g_bForm = True
4     g_sDate = ""
5     frmCalendar.Show_Cal txtToDate
6     If g_sDate <> "" Then
7         txtFromDate.Text = Format(FirstDayInMonth(CDate(g_sDate)), "dd-mmm-yyyy")
8     End If

Proc_Exit:
9     PopCallStack
10    Exit Sub

Err_Handler:
11    GlobalErrHandler
12    Resume Proc_Exit
End Sub

Private Sub txtToDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtCoopToDate_MouseDown"

3     g_bForm = True
4     g_sDate = ""
5     frmCalendar.Show_Cal txtFromDate
6     If g_sDate <> "" Then
7         txtToDate.Text = Format(LastDayInMonth(CDate(g_sDate)), "dd-mmm-yyyy")
8     End If

Proc_Exit:
9     PopCallStack
10    Exit Sub

Err_Handler:
11    GlobalErrHandler
12    Resume Proc_Exit
End Sub

Private Sub txtFromExtnDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtCoopFromDate_MouseDown"

3     g_bForm = True
4     g_sDate = ""
5     frmCalendar.Show_Cal txtToExtnDate
6     If g_sDate <> "" Then
7         txtFromExtnDate.Text = Format(FirstDayInMonth(CDate(g_sDate)), "dd-mmm-yyyy")
8     End If

Proc_Exit:
9     PopCallStack
10    Exit Sub

Err_Handler:
11    GlobalErrHandler
12    Resume Proc_Exit
End Sub

Private Sub txtToExtnDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtCoopToDate_MouseDown"

3     g_bForm = True
4     g_sDate = ""
5     frmCalendar.Show_Cal txtFromExtnDate
6     If g_sDate <> "" Then
7         txtToExtnDate.Text = Format(LastDayInMonth(CDate(g_sDate)), "dd-mmm-yyyy")
8     End If

Proc_Exit:
9     PopCallStack
10    Exit Sub

Err_Handler:
11    GlobalErrHandler
12    Resume Proc_Exit
End Sub
' Dates---------------------------------------------------End

' Save Record---------------------------------------------Start
' Submit for Approval
Private Sub cmdSubmitForApproval_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdSubmitForApproval_Click"

3     If ValidateInputs(Me) = True Then
4         Call SaveInputs("For Approval")
          
5         MsgBox "On Premise Contract " & txtRefNumber & " saved.", vbInformation, "Contract Saved"
6     End If

Proc_Exit:
7     PopCallStack
8     Exit Sub

Err_Handler:
9     GlobalErrHandler
10    Resume Proc_Exit
End Sub

' Save as Draft
Private Sub cmdSaveAsDraft_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdSaveAsDraft_Click"

3     Call SaveInputs("Draft")

4     MsgBox "On Premise Contract " & txtRefNumber & " saved as Draft.", vbInformation, "Contract Saved"

Proc_Exit:
5     PopCallStack
6     Exit Sub

Err_Handler:
7     GlobalErrHandler
8     Resume Proc_Exit
End Sub

Private Sub SaveInputs(strSaveType As String)
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|SaveInputs"

3     Select Case g_iAccessType
          Case enumUserPermission.OrdinaryUser
              ' Save only if Status is not yet approved for ordinary users
4             If CInt(lblStat.Caption) <> enumStatus.statApproved Then
5                 Call SaveRecordToDB(Me, strSaveType)
                  ' Hide Contract details
6                 If strSaveType <> "View" Then Call HideDetails
7             Else
8                 MsgBox "You cannot modify an already approved contract.", vbOKOnly, "Saving not Allowed"
9             End If
              
10        Case enumUserPermission.Admin, enumUserPermission.Manager
11            If CInt(lblStat.Caption) = enumStatus.statApproved And strSaveType <> "View" Then
12                If MsgBox("You are changing the Status of an already approved contract to """ & strSaveType & """. " & vbCrLf & vbCrLf & _
                            "Continue?", vbYesNo, "Confirm Status Change") = vbYes Then
13                    Call SaveRecordToDB(Me, strSaveType)
                      ' Hide Contract details
14                    Call HideDetails
15                End If
16            Else
17                Call SaveRecordToDB(Me, strSaveType)
                  ' Hide Contract details
18                If strSaveType <> "View" Then Call HideDetails
19            End If
              
20    End Select

Proc_Exit:
21    PopCallStack
22    Exit Sub

Err_Handler:
23    GlobalErrHandler
24    Resume Proc_Exit
End Sub
' Save Record---------------------------------------------End

Private Sub cmdDelete_Click()
      Dim intReply As Integer
      Dim qry As String

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdDelete_Click"

3     intReply = MsgBox("Delete the COOP with Ref# " & txtRefNumber & "?", vbYesNo, "Confirm Deletion")

4     If intReply = vbYes Then
5         If IsItemExistInTable(OP_MAIN_TBL, "RefNumber", txtRefNumber, "'") Then
              ' Set COOP record as deleted
6              qry = "UPDATE " & OP_MAIN_TBL & " " & _
                     "SET StatusID = " & enumStatus.statDeleted & ", " & _
                     "LastSyncDate = #" & ConvertLocalToGMT(Now) & "# " & _
                     "WHERE RefNumber = '" & txtRefNumber & "';"
7             cn.Execute qry
              
              ' Update last sync local date
8             qry = "UPDATE " & SYNC_DATE_TBL & " " & _
                    "SET LastSyncDate = #" & ConvertLocalToGMT(Now) & "# " & _
                    "WHERE ID = '" & GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "ID", "WinLoginName", UCase(Environ("UserName")), """") & "'"
9             cn.Execute qry

10            MsgBox "COOP '" & txtRefNumber & "' deleted successfully.", vbInformation
11        Else
              ' No existing Ref#
12            MsgBox "Reference number not existing. Nothing deleted.", vbExclamation
13        End If

          ' Hide Contract details
14        Call HideDetails
15    End If

Proc_Exit:
16    PopCallStack
17    Exit Sub

Err_Handler:
18    GlobalErrHandler
19    Resume Proc_Exit
End Sub

Private Sub cmdReplicate_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdReplicate_Click"
          
      ' Generate a new Ref#
3     txtRefNumber = GenerateRefNum

      ' Set Creator
4     cboCreator.Text = GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "[Name]", "WinLoginName", g_sLoginID, "'")

      ' Hide controls
5     cmdReplicate.Visible = False
6     cmdDelete.Visible = False

7     MsgBox "COOP replicated with the new Reference number.", vbInformation

Proc_Exit:
8     PopCallStack
9     Exit Sub

Err_Handler:
10    GlobalErrHandler
11    Resume Proc_Exit
End Sub

Private Sub cmdClose_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdClose_Click"

3     Call HideDetails

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

' Product selection---------------------------------------Start
Private Sub PopulateProductSelection()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|PopulateProductSelection"

3     cboProdType.List = Split(GetItemFromMappingTbl(SETTINGS_TBL, "Settings_Value", "Settings_Name", "Product_Types", """"), "|")
4     cboBrand.List = GetArrayList(AddSelectAll(PRODUCT_MAP_TBL, 2) & "SELECT DISTINCT BRAND_NAME, BRAND_CODE FROM " & PRODUCT_MAP_TBL & ";", True)
5     cboSubBrand.List = GetArrayList(AddSelectAll(PRODUCT_MAP_TBL, 2) & "SELECT DISTINCT SUB_BRAND_NAME, SUB_BRAND_CODE FROM " & PRODUCT_MAP_TBL & ";", True)
6     cboProdDescription.List = GetArrayList(AddSelectAll(PRODUCT_MAP_TBL, 4) & "SELECT DISTINCT PRODUCT_DESCRIPTION, PRODUCT_CODE, BOTTLE_SIZE, UNITS_PER_CASE FROM " & PRODUCT_MAP_TBL & ";", True)

Proc_Exit:
7     PopCallStack
8     Exit Sub

Err_Handler:
9     GlobalErrHandler
10    Resume Proc_Exit
End Sub

Private Sub cboBrand_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cboBrand_Change"

3     If cboBrand.Text = "(Select All)" Then cboBrand.Text = vbNullString
4     Call UpdateProductFilterList(cboBrand)

Proc_Exit:
5     PopCallStack
6     Exit Sub

Err_Handler:
7     GlobalErrHandler
8     Resume Proc_Exit
End Sub

Private Sub cboSubBrand_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cboSubBrand_Change"

3     If cboSubBrand.Text = "(Select All)" Then cboSubBrand.Text = vbNullString
4     Call UpdateProductFilterList(cboSubBrand)

Proc_Exit:
5     PopCallStack
6     Exit Sub

Err_Handler:
7     GlobalErrHandler
8     Resume Proc_Exit
End Sub

Private Sub cboProdDescription_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cboProdDescription_Change"

3     If cboProdDescription.Text = "(Select All)" Then cboProdDescription.Text = vbNullString
4     Call UpdateProductFilterList(cboProdDescription)

      ' Update Product code, Bottle size, Units per case, and Family(Spirit or Wine)
5     If cboProdDescription.ListIndex <> -1 Then
6         With cboProdDescription
7             txtProductCode.Text = .List(.ListIndex, 1)
8             txtBottleSize.Text = .List(.ListIndex, 2)
9             txtUnitsPerCase.Text = .List(.ListIndex, 3)
10        End With

11        txtFamily.Text = GetItemFromMappingTbl(PRODUCT_MAP_TBL, "FAMILY_NAME", "PRODUCT_CODE", txtProductCode.Text, """")
      '    If txtFamily.Text = "OTHER ALCOHOLIC BEVERAGES" Then
      '        txtFamily.Text = "Others"
      '    End If
12    Else
13        txtProductCode.Text = vbNullString
14        txtBottleSize.Text = vbNullString
15        txtUnitsPerCase.Text = vbNullString
16        txtFamily.Text = vbNullString
17    End If

      ' Re-calc Forecast Volume
18    txtContractedVolume = CalcForecastVolume

      ' Re-calc Forecast GSV
19    txtContractedGSV = CalcForecastGSV(txtProductCode, cboVarPrice.Text, SetEmptyValue(txtContractedCases, ZeroValue))

Proc_Exit:
20    PopCallStack
21    Exit Sub

Err_Handler:
22    GlobalErrHandler
23    Resume Proc_Exit
End Sub

' Dynamically change filter list
' Don't change the filters for calling control
Private Sub UpdateProductFilterList(ctrl As Control)
          ' This test is to prevent triggering another change event caused by ChangeBonusProductFilterList procedure
1         If blnFilterUpdating = False Then
2             blnFilterUpdating = True

3             If ctrl.Name <> "cboBrand" Or ctrl.Text = vbNullString Then _
                  Call ChangeProductFilterList("BRAND_NAME, BRAND_CODE", cboBrand)

4             If ctrl.Name <> "cboSubBrand" Or ctrl.Text = vbNullString Then _
                  Call ChangeProductFilterList("SUB_BRAND_NAME, SUB_BRAND_CODE", cboSubBrand)

5             If ctrl.Name <> "cboProdDescription" Or ctrl.Text = vbNullString Then _
                  Call ChangeProductFilterList("PRODUCT_DESCRIPTION, PRODUCT_CODE, BOTTLE_SIZE, UNITS_PER_CASE", cboProdDescription)

6             blnFilterUpdating = False
7         End If
End Sub

Private Sub ChangeProductFilterList(strCol As String, ctrl As Control)
          Dim qry As String
          Dim strWHERE As String
          Dim arr As Variant


1         qry = vbNullString
2         qry = qry & "SELECT DISTINCT " & strCol & " FROM " & PRODUCT_MAP_TBL & " "

          ' WHERE
3         strWHERE = vbNullString

4         If Len(cboBrand.Text) <> 0 And ctrl.Name <> "cboBrand" Then
5             strWHERE = strWHERE & IIf(Len(strWHERE) <> 0, "AND ", vbNullString)
6             strWHERE = strWHERE & "BRAND_CODE IN (""" & cboBrand.Value & """)" & vbCrLf
7         End If

8         If Len(cboSubBrand.Text) <> 0 And ctrl.Name <> "cboSubBrand" Then
9             strWHERE = strWHERE & IIf(Len(strWHERE) <> 0, "AND ", vbNullString)
10            strWHERE = strWHERE & "SUB_BRAND_CODE IN (""" & cboSubBrand.Value & """)" & vbCrLf
11        End If

12        If Len(cboProdDescription) <> 0 And ctrl.Name <> "cboProdDescription" Then
13            strWHERE = strWHERE & IIf(Len(strWHERE) <> 0, "AND ", vbNullString)
14            strWHERE = strWHERE & "PRODUCT_CODE IN (""" & cboProdDescription.Value & """)" & vbCrLf
15        End If

          ' Add Where clause
16        qry = qry & IIf(Len(strWHERE) <> 0, "WHERE " & strWHERE, "")

          ' Get list of filters
17        arr = GetArrayList(AddSelectAll(PRODUCT_MAP_TBL, UBound(Split(strCol, ",")) + 1) & qry, True)

          ' Check if it has returned any data
18        If IsArrayAllocated(arr) Then
19            ctrl.List = arr
20        End If

End Sub

Private Sub txtContractedCases_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtContractedCases_Change"

3     txtContractedVolume = CalcForecastVolume

4     txtContractedGSV = CalcForecastGSV(txtProductCode, cboVarPrice.Text, SetEmptyValue(txtContractedCases, ZeroValue))

Proc_Exit:
5     PopCallStack
6     Exit Sub

Err_Handler:
7     GlobalErrHandler
8     Resume Proc_Exit
End Sub

Private Function CalcForecastVolume() As Variant
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|CalcForecastVolume"

3     CalcForecastVolume = 0
4     If Not HasInvalidNumber(txtContractedCases) Then
5         CalcForecastVolume = SetEmptyValue(txtBottleSize.Text, ZeroValue) / 1000 * _
                               SetEmptyValue(txtUnitsPerCase.Text, ZeroValue) * _
                               SetEmptyValue(txtContractedCases, ZeroValue)
6     End If

7     If CalcForecastVolume = 0 Then CalcForecastVolume = vbNullString

Proc_Exit:
8     PopCallStack
9     Exit Function

Err_Handler:
10    GlobalErrHandler
11    Resume Proc_Exit
End Function

Private Function CalcForecastGSV(strProdCode As String, strVarPrice As String, dblForecastCases As Variant) As Variant
      Dim dblExcise As Double
      Dim dblVarPrice As Double

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|CalcForecastGSV"

3     CalcForecastGSV = 0
4     If Len(strProdCode) <> 0 And (IsNumeric(dblForecastCases)) Then
          ' Get Excise
5         dblExcise = SetEmptyValue(GetItemFromMappingTbl(EXCISE_MAP_TBL, "Excise", strWhereCondit:="ProductCode = """ & strProdCode & """ AND Start_Date <=#" & GetPromoDate(End_Date, Me) & "# AND End_Date >=#" & GetPromoDate(Start_Date, Me) & "#"), ZeroValue)
          
          ' Get Variable Price
6         dblVarPrice = SetEmptyValue(GetVariablePricing(strVarPrice, strProdCode), ZeroValue)
          
          ' Calculate Forecast GSV
7         CalcForecastGSV = (dblVarPrice - dblExcise) * dblForecastCases
8     End If

9     If CalcForecastGSV = 0 Then CalcForecastGSV = vbNullString

Proc_Exit:
10    PopCallStack
11    Exit Function

Err_Handler:
12    GlobalErrHandler
13    Resume Proc_Exit
End Function

Private Function CalcLUC(varVarPrice As Variant, varQAPerCase As Variant, varSize As Variant, varUnitsPerCase As Variant, varWet As Variant, _
                         varAdminAndFreight As Variant, varTTPerLtr As Variant, varTTGSV As Variant) As Variant
      Dim dblLUC As Double

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|CalcLUC"

3     dblLUC = 0
4     If IsNumeric(SetEmptyValue(varVarPrice, NullStrings)) And IsNumeric(SetEmptyValue(varQAPerCase, NullStrings)) And IsNumeric(SetEmptyValue(varUnitsPerCase, NullStrings)) Then
          'dblLUC = (((((varVarPrice - varQAPerCase) * varWet) + varAdminAndFreight) - (varTTPerLtr / ((varSize / 1000) * varUnitsPerCase)) - ((varTTGSV / 100) * (varVarPrice - varQAPerCase)))) / varUnitsPerCase
5         dblLUC = (((varVarPrice - varQAPerCase) * varWet) + varAdminAndFreight) / varUnitsPerCase
6     End If

7     If dblLUC = 0 Then
8         CalcLUC = vbNullString
9     Else
10        CalcLUC = dblLUC
11    End If

Proc_Exit:
12    PopCallStack
13    Exit Function

Err_Handler:
14    GlobalErrHandler
15    Resume Proc_Exit
End Function

Private Function CalcNIPPrice(varLUC As Variant, varSize As Variant) As Variant
      Dim dblNIPPrice As Double

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|CalcNIPPrice"

3     dblNIPPrice = 0
4     If IsNumeric(varLUC) And IsNumeric(varSize) And varSize <> 0 Then
5         dblNIPPrice = varLUC / (varSize / g_dblNIP_Const)
6     End If

7     If dblNIPPrice = 0 Then
8         CalcNIPPrice = vbNullString
9     Else
10        CalcNIPPrice = dblNIPPrice
11    End If

Proc_Exit:
12    PopCallStack
13    Exit Function

Err_Handler:
14    GlobalErrHandler
15    Resume Proc_Exit
End Function

Private Function GetKWI(dblGSV As Double, lstWS As MSForms.ListBox, strProductCode As String) As Variant
      Dim rs As ADODB.Recordset
      Dim qry As String
      Dim strPaidOn As String
      Dim strFamily As String
      Dim strWS As String
      Dim dblExcise As Double

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|GetKWI"

      ' Get Wholesaler
3     If LBXSelectCount(lstWS) > 1 Then
4         strWS = "ALM"
5     ElseIf LBXSelectCount(lstWS) = 1 Then
6         strWS = LBXSelectedItems(lstWS)(0)
7     Else    ' Direct no KWI
8         GetKWI = 0
9         GoTo Proc_Exit
10    End If

      ' Get Family
11    Select Case UCase(GetItemFromMappingTbl(PRODUCT_MAP_TBL, "FAMILY_NAME", "PRODUCT_CODE", strProductCode, """"))
          Case "SPIRITS"
12            strFamily = "Spirits"
13        Case "WINE"
14            strFamily = "Wine"
15        Case "OTHER ALCOHOLIC BEVERAGES"
16            strFamily = "Other"
17    End Select

      ' Get Paid_On and amount
18    Set rs = New ADODB.Recordset
19    qry = "SELECT Paid_On, Amount FROM " & KWI_MAP_TBL & " " & _
            "WHERE WS_Code = '" & strWS & "' AND Prod_Family = '" & strFamily & "' " & _
              "AND Start_Date <=#" & GetPromoDate(End_Date, Me) & "# AND End_Date >=#" & GetPromoDate(Start_Date, Me) & "#"
20    rs.Open qry, cn

21    If Not rs.EOF Then
          ' Calc KWI based on Paid_On
22        Select Case rs.Fields("Paid_On").Value
              Case "GSV"
23                GetKWI = dblGSV * rs.Fields("Amount").Value
24            Case "GSV+Excise"
25                dblExcise = SetEmptyValue(GetItemFromMappingTbl(EXCISE_MAP_TBL, "Excise", strWhereCondit:="ProductCode = """ & strProductCode & """ AND Start_Date <=#" & GetPromoDate(End_Date, Me) & "# AND End_Date >=#" & GetPromoDate(Start_Date, Me) & "#"), ZeroValue)
26                GetKWI = (dblGSV + dblExcise) * rs.Fields("Amount").Value
27        End Select
28    End If
29    Call CloseRecordset(rs, True)

Proc_Exit:
30    PopCallStack
31    Exit Function

Err_Handler:
32    GlobalErrHandler
33    Resume Proc_Exit
End Function

Private Function GetCOP(varGSV As Variant, varQA3 As Variant, lstWS As MSForms.ListBox, strProductCode As String) As Variant
      Dim rs As ADODB.Recordset
      Dim qry As String
      Dim strPaidOn As String
      Dim strFamily As String
      Dim strWS As String
      Dim dblExcise As Double

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|GetCOP"

3     GetCOP = 0

      ' Check if Wholesaler is ALM
4     If cboRouteToMarket = "Indirect" And LBXSelectCount(lstWholesaler) <> 0 Then
5         If InStr(1, GetIN_List(LBXSelectedItems(lstWholesaler, 0), vbNullString, "|"), "ALM") = 0 Then
6             GoTo Proc_Exit
7         End If
8     Else
9         GoTo Proc_Exit
10    End If

      ' Get Family
11    strFamily = UCase(GetItemFromMappingTbl(PRODUCT_MAP_TBL, "FAMILY_NAME", "PRODUCT_CODE", strProductCode, """"))

      '' Identify if MatchCode is COP
      'If SetEmptyValue(GetItemFromMappingTbl(MATCHCODE_ALMCUSTCODE_MAP_TBL, "ALM_Cust_Code", "MatchCode", strFamily, """"), NullStrings) = vbNullString Then
      '    GoTo Proc_Exit
      'End If

      ' Get Paid_On and Amount
12    Set rs = New ADODB.Recordset
13    qry = "SELECT Paid_On, Amount FROM " & COP_TERMS_MAP_TBL & " " & _
            "WHERE Family = '" & strFamily & "' " & _
              "AND Start_Date <=#" & GetPromoDate(End_Date, Me) & "# AND End_Date >=#" & GetPromoDate(Start_Date, Me) & "#"
14    rs.Open qry, cn

15    If Not rs.EOF Then
16        Select Case rs.Fields("Paid_On").Value
              Case "GSV"
17                GetCOP = (varGSV - varQA3) * rs.Fields("Amount").Value
18            Case "GSV+Excise"
19                dblExcise = SetEmptyValue(GetItemFromMappingTbl(EXCISE_MAP_TBL, "Excise", strWhereCondit:="ProductCode = """ & strProductCode & """ AND Start_Date <=#" & GetPromoDate(End_Date, Me) & "# AND End_Date >=#" & GetPromoDate(Start_Date, Me) & "#"), ZeroValue)
20                GetCOP = ((varGSV - varQA3) + dblExcise) * rs.Fields("Amount").Value
21        End Select
22    End If
23    Call CloseRecordset(rs, True)

Proc_Exit:
24    PopCallStack
25    Exit Function

Err_Handler:
26    GlobalErrHandler
27    Resume Proc_Exit
End Function

Private Function GetVariablePricing(strVarPrice As String, strProdCode As String) As Variant
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|GetVariablePricing"

3     GetVariablePricing = 0

4     GetVariablePricing = SetEmptyValue(GetItemFromMappingTbl(PRICING_MAP_TBL, "[" & strVarPrice & "]", strWhereCondit:="ProductCode = """ & strProdCode & """ AND Start_Date <=#" & GetPromoDate(End_Date, Me) & "# AND End_Date >=#" & GetPromoDate(Start_Date, Me) & "#"), ZeroValue)

5     If GetVariablePricing = 0 Then GetVariablePricing = vbNullString

Proc_Exit:
6     PopCallStack
7     Exit Function

Err_Handler:
8     GlobalErrHandler
9     Resume Proc_Exit
End Function

Private Sub cmdClearInputsProd_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdClearInputsProd_Click"

3     cboProdType.Value = vbNullString
4     cboBrand.Text = vbNullString
5     cboSubBrand.Text = vbNullString
6     cboProdDescription.Text = vbNullString
7     Call PopulateProductSelection
8     txtProductCode = vbNullString
9     txtBottleSize = vbNullString
10    txtUnitsPerCase = vbNullString
11    txtContractedCases = vbNullString
12    txtContractedVolume = vbNullString
13    txtContractedGSV = vbNullString

Proc_Exit:
14    PopCallStack
15    Exit Sub

Err_Handler:
16    GlobalErrHandler
17    Resume Proc_Exit
End Sub
' Product selection---------------------------------------End

' Add Product---------------------------------------------Start
Private Sub cmdAddProduct_Click()
      Dim arrProductInfo As Variant
      Dim x As Integer, y As Integer
      Dim arrData As Variant

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdAddProduct_Click"

      ' Validate inputs
      ' Product Type
3     If HasEmptyValues(cboProdType, "Product Type") Then GoTo Proc_Exit
      ' Product Description
4     If HasEmptyValues(cboProdDescription, "Product Description") Then GoTo Proc_Exit
      ' Product Code
5     If HasEmptyValues(txtProductCode, strName:="Product Code") Then GoTo Proc_Exit
      ' Bottle Size
6     If HasInvalidNumber(txtBottleSize, strName:="Bottle Size") Then GoTo Proc_Exit
      ' Units Per Case
7     If HasInvalidNumber(txtUnitsPerCase, strName:="Units Per Case") Then GoTo Proc_Exit
      ' Forecast Cases
8     If HasInvalidNumber(txtContractedCases, strName:="Contractual Cases") Then GoTo Proc_Exit

      ' Get Product info
9     arrProductInfo = GetProductInfo(cboProdType, cboBrand, cboSubBrand, cboProdDescription, _
                       ", '" & txtContractedCases & "', '" & txtContractedVolume & "', '" & RoundNum(txtContractedVolume, 0) & "', '" & txtContractedGSV & _
                       "', '" & RoundNum(txtContractedGSV, 0) & "', '" & txtFamily & "'")

      ' Add to list
10    If IsArrayAllocated(arrProductInfo) Then
11        With lstProducts
12            If .ListCount = 0 Then
13                ReDim arrData(0, UBound(arrProductInfo, 2))
14            Else
15                ReDim arrData(.ListCount, .ColumnCount - 1)
                  
16                For x = 0 To .ListCount - 1
17                    For y = 0 To .ColumnCount - 1
18                        arrData(x, y) = .List(x, y)
19                    Next y
20                Next x
21            End If
                      
              ' Add the new input to the array
22            For y = 0 To UBound(arrProductInfo, 2)
23                arrData(UBound(arrData), y) = arrProductInfo(0, y)
24            Next y
              
              ' Refresh listbox
25            .Clear
26            .List = arrData
27        End With

28        Call AddToTradingTerms(arrProductInfo)
29        Call AddToQA3(arrProductInfo)
30    End If

      ' Clear inputs
31    cmdClearInputsProd_Click

      ' Clear array data
32    Erase arrData

Proc_Exit:
33    PopCallStack
34    Exit Sub

Err_Handler:
35    GlobalErrHandler
End Sub

Private Function GetProductInfo(strProdType As String, cboBrd As MSForms.ComboBox, cboSubBrd As MSForms.ComboBox, cboProdDesc As MSForms.ComboBox, strOthers As String) As Variant
      Dim qry As String
      Dim strWHERE As String
      Dim arr As Variant

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|GetProductInfo"

3     qry = vbNullString
4     qry = qry & "SELECT DISTINCT '" & strProdType & "', BRAND_NAME, BRAND_CODE, SUB_BRAND_NAME, SUB_BRAND_CODE, PRODUCT_DESCRIPTION, PRODUCT_CODE, BOTTLE_SIZE, UNITS_PER_CASE " & strOthers & "FROM " & PRODUCT_MAP_TBL & " "

      ' WHERE
5     strWHERE = vbNullString

6     If Len(cboBrd.Value) <> 0 Then
7         strWHERE = strWHERE & IIf(Len(strWHERE) <> 0, "AND ", vbNullString)
8         strWHERE = strWHERE & "BRAND_CODE IN (""" & cboBrd.Value & """)" & vbCrLf
9     End If

10    If Len(cboSubBrd.Value) <> 0 Then
11        strWHERE = strWHERE & IIf(Len(strWHERE) <> 0, "AND ", vbNullString)
12        strWHERE = strWHERE & "SUB_BRAND_CODE IN (""" & cboSubBrd.Value & """)" & vbCrLf
13    End If

14    If Len(cboProdDesc.Value) <> 0 Then
15        strWHERE = strWHERE & IIf(Len(strWHERE) <> 0, "AND ", vbNullString)
16        strWHERE = strWHERE & "PRODUCT_CODE IN (""" & cboProdDesc.Value & """)" & vbCrLf
17    End If

      ' Add Where clause
18    qry = qry & IIf(Len(strWHERE) <> 0, "WHERE " & strWHERE, "")

      ' Get the filters list
19    arr = GetArrayList(qry, True)

      ' Check if it has returned any data
20    If IsArrayAllocated(arr) Then
21        GetProductInfo = arr
22    End If

Proc_Exit:
23    PopCallStack
24    Exit Function

Err_Handler:
25    GlobalErrHandler
26    Resume Proc_Exit
          
End Function

Private Sub AddToTradingTerms(arrProd As Variant) ' strProdDesc As String, strProdCode As String, strLitres As String, strGSV As String)
      Dim x As Integer, y As Integer
      Dim arrData As Variant

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|AddToTradingTerms"

3     With lstTrdTerms
4         If .ListCount = 0 Then
5             ReDim arrData(0, .ColumnCount - 1)
6         Else
7             ReDim arrData(.ListCount, .ColumnCount - 1)
              
8             For x = 0 To .ListCount - 1
9                 For y = 0 To .ColumnCount - 1
10                    arrData(x, y) = .List(x, y)
11                Next y
12            Next x
13        End If
                  
          ' Add the new input to the array
14        arrData(UBound(arrData), TTList_ProdType) = arrProd(0, ProdList_ProdType)
15        arrData(UBound(arrData), TTList_Brand) = arrProd(0, ProdList_Brand)
16        arrData(UBound(arrData), TTList_ProdDesc) = arrProd(0, ProdList_ProdDesc)
17        arrData(UBound(arrData), TTList_ProdCode) = arrProd(0, ProdList_ProdCode)
18        arrData(UBound(arrData), TTList_ContractVol) = arrProd(0, ProdList_ContractVolRoundoff)
19        arrData(UBound(arrData), TTList_ContractGSV) = arrProd(0, ProdList_ContractGSVRoundoff)
          
          ' Refresh listbox
20        .Clear
21        .List = arrData
22    End With

Proc_Exit:
23    PopCallStack
24    Exit Sub

Err_Handler:
25    GlobalErrHandler
26    Resume Proc_Exit
End Sub

Private Sub AddToQA3(arrProd As Variant) 'strProdDesc As String, strProdCode As String, strLitres As String, strGSV As String, strFamily As String)
      Dim x As Integer, y As Integer
      Dim arrData As Variant

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|AddToQA3"

3     With lstQA3
4         If .ListCount = 0 Then
5             ReDim arrData(0, .ColumnCount - 1)
6         Else
7             ReDim arrData(.ListCount, .ColumnCount - 1)
              
8             For x = 0 To .ListCount - 1
9                 For y = 0 To .ColumnCount - 1
10                    arrData(x, y) = .List(x, y)
11                Next y
12            Next x
13        End If
                  
          ' Add the new input to the array
14        arrData(UBound(arrData), QA3List_ProdType) = arrProd(0, ProdList_ProdType)
15        arrData(UBound(arrData), QA3List_Brand) = arrProd(0, ProdList_Brand)
16        arrData(UBound(arrData), QA3List_ProdDesc) = arrProd(0, ProdList_ProdDesc)
17        arrData(UBound(arrData), QA3List_ProdCode) = arrProd(0, ProdList_ProdCode)
18        arrData(UBound(arrData), QA3List_ContractVol) = arrProd(0, ProdList_ContractVolRoundoff)
19        arrData(UBound(arrData), QA3List_ContractGSV) = arrProd(0, ProdList_ContractGSVRoundoff)
20        If cboRouteToMarket = "Direct" Then
21            arrData(UBound(arrData), QA3List_DirectPrice) = GetVariablePricing(cboVarPrice.Text, CStr(arrProd(0, ProdList_ProdCode)))
22            arrData(UBound(arrData), QA3List_DirectPriceRoundoff) = RoundNum(arrData(UBound(arrData), QA3List_DirectPrice), 2)
23        End If
24        arrData(UBound(arrData), QA3List_KWI) = GetKWI(SetEmptyValue(arrProd(0, ProdList_ContractGSV), ZeroValue), lstWholesaler, CStr(arrProd(0, ProdList_ProdCode)))
25        arrData(UBound(arrData), QA3List_Family) = arrProd(0, ProdList_Family)
          
          ' Refresh listbox
26        .Clear
27        .List = arrData
28    End With

Proc_Exit:
29    PopCallStack
30    Exit Sub

Err_Handler:
31    GlobalErrHandler
32    Resume Proc_Exit
End Sub
' Add Product---------------------------------------------End

' Delete Product------------------------------------------Start
Private Sub cmdDeleteProduct_Click()
      Dim arr As Variant

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdDeleteProduct_Click"

3     If LBXSelectCount(lstProducts) <> 0 Then
4         If MsgBox("Are you sure you want to delete the selected record(s)?", vbYesNo, "Confirm Deletion") = vbYes Then
              ' Get selected index
5             arr = LBXSelectedIndexes(lstProducts)
                      
              ' Delete from both Products, Trading Terms, and QA3 tabs
6             Call DelSelProduct(arr, lstProducts)
7             Call DelSelProduct(arr, lstTrdTerms)
8             Call DelSelProduct(arr, lstQA3)
9         End If
10    End If

Proc_Exit:
11    PopCallStack
12    Exit Sub

Err_Handler:
13    GlobalErrHandler
14    Resume Proc_Exit
End Sub

Private Sub DelSelProduct(arrSelected As Variant, lst As MSForms.ListBox)
      Dim i As Long, x As Integer, y As Integer
      Dim arrData As Variant

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|DelSelProduct"

3     With lst
4         If .ListCount - UBound(arrSelected) - 2 <= -1 Then
5             .Clear
6         Else
7             If IsArrayAllocated(arrSelected) Then
                  
8                 ReDim arrData(.ListCount - UBound(arrSelected) - 2, .ColumnCount - 1)
                  
9                 i = 0
10                For x = 0 To .ListCount - 1
11                    If IsInArray(arrSelected, x) = False Then
12                        For y = 0 To .ColumnCount - 1
13                            arrData(i, y) = .List(x, y)
14                        Next y
15                        i = i + 1
16                    End If
17                Next x
                  
                  ' Refresh listbox
18                .Clear
19                .List = arrData
                  
20                Erase arrData
21            End If
22        End If
23    End With

Proc_Exit:
24    PopCallStack
25    Exit Sub

Err_Handler:
26    GlobalErrHandler
27    Resume Proc_Exit
End Sub
' Delete Product------------------------------------------End

' Edit Product--------------------------------------------Start
Private Sub cmdEditProduct_Click()
      Dim idx As Integer

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdEditProduct_Click"

      ' Check if there are record selected
3     Select Case LBXSelectCount(lstProducts)
          Case Is = 1 ' Allow edit if user selects only 1
4             Select Case cmdEditProduct.Caption
                  Case "Edit"
                      ' Change button caption
5                     cmdEditProduct.Caption = "Update"
                      
                      ' Clear inputs
6                     Call cmdClearInputsProd_Click
                      
                      ' Disable other controls
7                     cmdClearInputsProd.Enabled = False
8                     cmdAddProduct.Enabled = False
9                     cmdDeleteProduct.Enabled = False
10                    fraRecordCommands.Enabled = False
11                    lstProducts.Locked = True
12                    mpgeOP.Pages(0).Enabled = False
13                    mpgeOP.Pages(2).Enabled = False
14                    mpgeOP.Pages(3).Enabled = False
15                    mpgeOP.Pages(4).Enabled = False
16                    mpgeOP.Pages(5).Enabled = False
                      
                      ' Disable product info
17                    cboBrand.Enabled = False
18                    cboSubBrand.Enabled = False
19                    cboProdDescription.Enabled = False

                      ' Populate fields
20                    With lstProducts
21                        idx = .ListIndex
                          
22                        cboProdType.Text = .List(idx, ProdList_ProdType)
23                        cboBrand.Text = .List(idx, ProdList_Brand)
24                        cboSubBrand.Text = .List(idx, ProdList_Subbrand)
25                        cboProdDescription.Text = .List(idx, ProdList_ProdDesc)
26                        txtContractedCases.Text = .List(idx, ProdList_ContractCases)
27                    End With
                      
28                Case "Update"
                      ' Validate inputs
29                    If HasInvalidNumber(txtContractedCases, strName:="Forecast Cases") Then GoTo Proc_Exit
30                    If HasInvalidNumber(txtContractedVolume, strName:="Forecast Volume") Then GoTo Proc_Exit
31                    If HasInvalidNumber(txtContractedGSV, strName:="Forecast GSV") Then GoTo Proc_Exit
                                     
32                    idx = lstProducts.ListIndex
                      
                      ' Update Products list
33                    With lstProducts
34                        .List(idx, ProdList_ProdType) = cboProdType.Text
35                        .List(idx, ProdList_ContractCases) = txtContractedCases.Text
36                        .List(idx, ProdList_ContractVol) = txtContractedVolume.Text
37                        .List(idx, ProdList_ContractVolRoundoff) = RoundNum(txtContractedVolume.Text, 0)
38                        .List(idx, ProdList_ContractGSV) = txtContractedGSV.Text
39                        .List(idx, ProdList_ContractGSVRoundoff) = RoundNum(txtContractedGSV.Text, 0)
40                    End With
                      
                      ' Update QA3 list
41                    With lstQA3
42                        .List(idx, QA3List_ContractVol) = RoundNum(txtContractedVolume.Text, 0)
43                        .List(idx, QA3List_ContractGSV) = RoundNum(txtContractedGSV.Text, 0)
44                        .List(idx, QA3List_KWI) = GetKWI(txtContractedGSV.Text, lstWholesaler, txtProductCode)
45                        .List(idx, QA3List_COP) = GetCOP(txtContractedGSV.Text, .List(idx, QA3List_QA3), lstWholesaler, txtProductCode)
46                    End With
                  
                      ' Update Trading Terms list
47                    With lstTrdTerms
48                        .List(idx, TTList_ContractVol) = RoundNum(txtContractedVolume.Text, 0)
49                        .List(idx, TTList_ContractGSV) = RoundNum(txtContractedGSV.Text, 0)
                          
                          ' Recalc Terms
50                        .List(idx, TTList_StandardTerm) = CalcTradingTerms(.List(idx, TTList_TTLtr), lstProducts.List(idx, ProdList_ContractVol), .List(idx, TTList_TTGSV), lstProducts.List(idx, ProdList_ContractGSV), lstQA3.List(idx, QA3List_QA3))
51                        .List(idx, TTList_AddnlTerm) = CalcAdditionalTerms(.List(idx, TTList_TTLtr), .List(idx, TTList_TTMaxLtr), lstProducts.List(idx, ProdList_ContractVol), .List(idx, TTList_TTGSV), .List(idx, TTList_TTMaxGSV), lstProducts.List(idx, ProdList_ContractGSV), lstQA3.List(idx, QA3List_QA3))
52                        .List(idx, TTList_BannerTerm) = CalcBannerTerms(txtTTBannerGSV, txtTTBannerGSVlessQA3, lstProducts.List(idx, ProdList_ContractGSV), lstQA3.List(idx, QA3List_QA3))
53                    End With
                    
                      ' Change button caption
54                    cmdEditProduct.Caption = "Edit"
                      
                      ' Clear inputs
55                    Call cmdClearInputsProd_Click
                      
                      ' Enable back the controls
56                    cmdClearInputsProd.Enabled = True
57                    cmdAddProduct.Enabled = True
58                    cmdDeleteProduct.Enabled = True
59                    fraRecordCommands.Enabled = True
60                    lstProducts.Locked = False
61                    mpgeOP.Pages(0).Enabled = True
62                    mpgeOP.Pages(2).Enabled = True
63                    mpgeOP.Pages(3).Enabled = True
64                    mpgeOP.Pages(4).Enabled = True
65                    mpgeOP.Pages(5).Enabled = True
                      
                      ' Enable product info
66                    cboBrand.Enabled = True
67                    cboSubBrand.Enabled = True
68                    cboProdDescription.Enabled = True
69            End Select

70        Case Is > 1 ' Don't allow edit if user selects multiple records
71            MsgBox "Please select only one record to edit.", vbExclamation

72    End Select

Proc_Exit:
73    PopCallStack
74    Exit Sub

Err_Handler:
75    GlobalErrHandler
76    Resume Proc_Exit
          
End Sub
' Edit Product--------------------------------------------End

' Edit QA3------------------------------------------------Start
Private Sub cmdClearInputsQA3_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdClearInputsQA3_Click"

3     txtQA3ProdType = vbNullString
4     txtQA3ProdBrand = vbNullString
5     txtQA3ProdDesc = vbNullString
6     txtQA3Litres = vbNullString
7     txtQA3GSV = vbNullString
8     txtDirectPrice = vbNullString
9     txtWholesalerPrice = vbNullString
10    txtQA3PerCaseUser = vbNullString
11    txtNIPOrLUCAuto = vbNullString
12    txtNIPOrLUCUser = vbNullString
13    txtQA3PerCaseAuto = vbNullString
14    txtProdUnitsPerCase = vbNullString
15    txtProdBottleSize = vbNullString
16    txtTrdTermsPerLtr = vbNullString
17    txtTrdTermsGSV = vbNullString
18    txtKWI = vbNullString
19    txtCOP = vbNullString
20    txtQA3Family = vbNullString

Proc_Exit:
21    PopCallStack
22    Exit Sub

Err_Handler:
23    GlobalErrHandler
24    Resume Proc_Exit
End Sub

Private Sub cmdEditQA3_Click()
      Dim idx As Integer

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdEditQA3_Click"

      ' Check if there are record selected
3     If LBXSelectCount(lstQA3) = 1 Then
4         Select Case cmdEditQA3.Caption
              Case "Edit"
                  ' Change button caption
5                 cmdEditQA3.Caption = "Update"
                  
                  ' Clear inputs
6                 Call cmdClearInputsQA3_Click
                  
                  ' Disable other controls
7                 cmdClearInputsQA3.Enabled = False
8                 fraRecordCommands.Enabled = False
9                 lstQA3.Locked = True
10                mpgeOP.Pages(0).Enabled = False
11                mpgeOP.Pages(1).Enabled = False
12                mpgeOP.Pages(2).Enabled = False
13                mpgeOP.Pages(4).Enabled = False
14                mpgeOP.Pages(5).Enabled = False

                  ' Populate fields
15                With lstQA3
16                    idx = .ListIndex
                      
17                    txtQA3ProdType = .List(idx, QA3List_ProdType)
18                    txtQA3ProdBrand = .List(idx, QA3List_Brand)
19                    txtQA3ProdDesc = .List(idx, QA3List_ProdDesc)
20                    txtQA3Litres = .List(idx, QA3List_ContractVol)
21                    txtQA3GSV = .List(idx, QA3List_ContractGSV)
22                    txtQA3Family.Text = .List(idx, QA3List_Family)
23                    txtProdUnitsPerCase.Text = lstProducts.List(idx, ProdList_UnitsPerCase)
24                    txtProdBottleSize.Text = lstProducts.List(idx, ProdList_BottleSize)
25                    txtTrdTermsPerLtr.Text = lstTrdTerms.List(idx, TTList_TTLtr)
26                    txtTrdTermsGSV.Text = lstTrdTerms.List(idx, TTList_TTGSV)
27                    txtDirectPrice.Text = .List(idx, QA3List_DirectPrice)
28                    txtWholesalerPrice.Text = .List(idx, QA3List_WSPrice)
29                    txtQA3PerCaseUser.Text = .List(idx, QA3List_QA3Input)
30                    txtNIPOrLUCUser.Text = .List(idx, QA3List_NipOrLUCInput)
31                End With
                  
32            Case "Update"
                  ' Validate inputs
                 
                  ' Direct or Wholesaler price
33                Select Case cboRouteToMarket.Value
                      Case "Direct"
34                        If HasInvalidNumber(txtDirectPrice, strName:="Direct Price") Then GoTo Proc_Exit
35                    Case "Indirect"
36                        If HasInvalidNumber(txtWholesalerPrice, strName:="Wholesaler Price") Then GoTo Proc_Exit
37                End Select
                  
                  ' User Defined QA3 or NIP/LUC
38                If Len(txtQA3PerCaseUser) <> 0 And Len(txtNIPOrLUCUser) <> 0 Then
39                    MsgBox "Only one of User Def QA3 or User Def NIP/LUC should exist.", vbExclamation
40                    GoTo Proc_Exit
41                ElseIf Len(txtQA3PerCaseUser) = 0 And Len(txtNIPOrLUCUser) = 0 Then
42                    MsgBox "Either one of User Def QA3 or User Def NIP/LUC should be inputted.", vbExclamation
43                    GoTo Proc_Exit
44                Else
45                    If Len(txtQA3PerCaseUser) <> 0 Then
46                        txtNIPOrLUCUser = vbNullString
47                        txtQA3PerCaseAuto = vbNullString
48                        If HasInvalidNumber(txtQA3PerCaseUser, strName:="QA3 Per Case") Then GoTo Proc_Exit
49                    ElseIf Len(txtNIPOrLUCUser) <> 0 Then
50                        txtQA3PerCaseUser = vbNullString
51                        txtNIPOrLUCAuto = vbNullString
52                        If HasInvalidNumber(txtNIPOrLUCUser, strName:="User Defined NIP/LUC") Then GoTo Proc_Exit
53                    End If
54                End If
                  
                  ' Update QA3 amounts
55                With lstQA3
                      ' Get index of selected item
56                    idx = .ListIndex
                      
57                    .List(idx, QA3List_DirectPrice) = txtDirectPrice.Text
58                    .List(idx, QA3List_DirectPriceRoundoff) = RoundNum(txtDirectPrice.Text, 2)
59                    .List(idx, QA3List_WSPrice) = txtWholesalerPrice.Text
60                    .List(idx, QA3List_WSPriceRoundoff) = RoundNum(txtWholesalerPrice.Text, 2)
61                    .List(idx, QA3List_QA3Input) = txtQA3PerCaseUser.Text
62                    .List(idx, QA3List_QA3InputRoundoff) = RoundNum(txtQA3PerCaseUser.Text, 2)
63                    .List(idx, QA3List_NipOrLUCAuto) = txtNIPOrLUCAuto.Text
64                    .List(idx, QA3List_NipOrLUCAutoRoundoff) = RoundNum(txtNIPOrLUCAuto.Text, 2)
65                    .List(idx, QA3List_NipOrLUCInput) = txtNIPOrLUCUser.Text
66                    .List(idx, QA3List_NipOrLUCInputRoundoff) = RoundNum(txtNIPOrLUCUser.Text, 2)
67                    .List(idx, QA3List_QA3Auto) = txtQA3PerCaseAuto.Text
68                    .List(idx, QA3List_QA3AutoRoundoff) = RoundNum(txtQA3PerCaseAuto.Text, 2)
69                    .List(idx, QA3List_QA3) = CalcQA3(SetEmptyValue(.List(idx, QA3List_QA3Input) + .List(idx, QA3List_QA3Auto), ZeroValue), lstProducts.List(idx, ProdList_ContractVol), lstProducts.List(idx, ProdList_BottleSize), lstProducts.List(idx, ProdList_UnitsPerCase))
70                    .List(idx, QA3List_QA3Roundoff) = RoundNum(.List(idx, QA3List_QA3), 2)
71                    .List(idx, QA3List_COP) = GetCOP(SetEmptyValue(lstProducts.List(idx, ProdList_ContractGSV), ZeroValue), SetEmptyValue(.List(idx, QA3List_QA3), ZeroValue), lstWholesaler, .List(idx, QA3List_ProdCode))
72                    .List(idx, QA3List_COPRoundoff) = RoundNum(.List(idx, QA3List_COP), 2)

73                End With
                
                  ' Update Trading Terms amounts
74                With lstTrdTerms
75                    .List(idx, TTList_StandardTerm) = CalcTradingTerms(.List(idx, TTList_TTLtr), lstProducts.List(idx, ProdList_ContractVol), .List(idx, TTList_TTGSV), lstProducts.List(idx, ProdList_ContractGSV), lstQA3.List(idx, QA3List_QA3))
76                    .List(idx, TTList_AddnlTerm) = CalcAdditionalTerms(.List(idx, TTList_TTLtr), .List(idx, TTList_TTMaxLtr), lstProducts.List(idx, ProdList_ContractVol), .List(idx, TTList_TTGSV), .List(idx, TTList_TTMaxGSV), lstProducts.List(idx, ProdList_ContractGSV), lstQA3.List(idx, QA3List_QA3))
77                    .List(idx, TTList_BannerTerm) = CalcBannerTerms(txtTTBannerGSV, txtTTBannerGSVlessQA3, lstProducts.List(idx, ProdList_ContractGSV), lstQA3.List(idx, QA3List_QA3))
78                End With
                  
                  ' Change button caption
79                cmdEditQA3.Caption = "Edit"
                  
                  ' Clear inputs
80                Call cmdClearInputsQA3_Click
                  
                  ' Enable back the controls
81                cmdClearInputsQA3.Enabled = True
82                fraRecordCommands.Enabled = True
83                lstQA3.Locked = False
84                mpgeOP.Pages(0).Enabled = True
85                mpgeOP.Pages(1).Enabled = True
86                mpgeOP.Pages(2).Enabled = True
87                mpgeOP.Pages(4).Enabled = True
88                mpgeOP.Pages(5).Enabled = True

89        End Select

90    End If

Proc_Exit:
91    PopCallStack
92    Exit Sub

Err_Handler:
93    GlobalErrHandler
94    Resume Proc_Exit

End Sub
' Edit QA3------------------------------------------------End

' View PEM Reports----------------------------------------Start
Private Sub cmdViewPEMReport_Click()
      Dim wb As Workbook

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdViewPEMReport_Click"

      ' Set start of transactions so that we can rollback any changes to the database
3     cn.BeginTrans

      ' View only if status is submitted for approval
4     If ValidateInputs(Me) = True Then
          ' Save inputs in database
5         Call SaveInputs("View")

6         Set wb = CreateViewWorkbook(Array(PEM_TEMP_SHEET, PEM_SUMM_TEMP_SHEET))
          
          ' Create view
7         Call CreatePEMReport(wb, Me)
8         Call CreateSummaryReport(wb, Me)
          
          ' Delete templates
9         Call DeleteViewTemplates(wb, Array(PEM_TEMP_SHEET, PEM_SUMM_TEMP_SHEET))
          
10        Set wb = Nothing
11    End If


Proc_Exit:
12    cn.RollbackTrans

13    PopCallStack
14    Exit Sub

Err_Handler:
15    GlobalErrHandler
16    Resume Proc_Exit
End Sub
' View PEM Reports----------------------------------------End

' View Deal Sheet Reports---------------------------------Start
Private Sub cmdDealSheets_Click()
      Dim wb As Workbook
      Dim i As Integer
      Dim arrWS As Variant

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdDealSheets_Click"

      ' Set start of transactions so that we can rollback any changes to the database
3     cn.BeginTrans

      ' View only if status is submitted for approval
4     If ValidateInputs(Me) = True Then
          ' Save inputs in database
5         Call SaveInputs("View")
             
          ' Deal Sheets
6         If cboRouteToMarket = "Indirect" And LBXSelectCount(lstWholesaler) <> 0 Then
7             Set wb = CreateViewWorkbook(Array(ALM_DEAL_TEMP_SHEET, STANDARD_DEAL_TEMP_SHEET))
              
8             arrWS = LBXSelectedItems(lstWholesaler)
9             For i = 0 To UBound(arrWS)
10                If arrWS(i) = "ALM" Then
                      ' Create ALM deal sheet
11                    Call CreateALMDealSheet(wb, Me)
12                Else
                      ' Create Standard deal sheet
13                    Call CreateStandardDealSheet(wb, Me, CStr(arrWS(i)))
14                End If
15            Next i
              
16            Call DeleteViewTemplates(wb, Array(ALM_DEAL_TEMP_SHEET, STANDARD_DEAL_TEMP_SHEET))
17            Set wb = Nothing
18        End If
          
          ' E1 upload Sheet
          ' View only if its Direct and in Outlet level
19        If cboRouteToMarket.Value = "Direct" And cboContractLevel.Value = "OP Outlet Level" Then
20            Set wb = CreateViewWorkbook(Array(E1_UPLOAD_TEMP_SHEET))
              
21            Call CreateE1UploadSheet(wb, Me)
              
22            Call DeleteViewTemplates(wb, Array(E1_UPLOAD_TEMP_SHEET))
23            Set wb = Nothing
24        End If
25    End If

Proc_Exit:
26    cn.RollbackTrans

27    PopCallStack
28    Exit Sub

Err_Handler:
29    GlobalErrHandler
30    Resume Proc_Exit
End Sub
' View Deal Sheet Reports---------------------------------End

' View Word Report----------------------------------------Start
Private Sub cmdWordReport_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdWordReport_Click"

      ' View only if status is submitted for approval
3     If ValidateInputs(Me) = True Then
4         Call GenerateWordDocs(Me)
5     End If

Proc_Exit:
6     PopCallStack
7     Exit Sub

Err_Handler:
8     GlobalErrHandler
9     Resume Proc_Exit
End Sub
' View Word Report----------------------------------------End

Private Sub lstContractLevelCode_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|lstContractLevelCode_Change"

3     txtOutletOrGroupName.Locked = True

4     If cboContractLevel = "OP Banner Region" Or cboContractLevel = "OP Banner" Then
5         txtOutletOrGroupName = lstContractLevelCode.List(LBXSelectedIndexes(lstContractLevelCode)(0), 0)
6     ElseIf LBXSelectCount(lstContractLevelCode) = 1 And cboContractLevel = "OP Outlet Level" Then
7         If CStr(LBXSelectedItems(lstContractLevelCode)(0)) = "Opportunity Outlet" Then
8             txtOutletOrGroupName = vbNullString
9             txtOutletOrGroupName.Locked = False
10        Else
11            txtOutletOrGroupName = lstContractLevelCode.List(LBXSelectedIndexes(lstContractLevelCode)(0), 0)
12        End If
13    Else
14        txtOutletOrGroupName = vbNullString
15        txtOutletOrGroupName.Locked = False
16    End If

Proc_Exit:
17    PopCallStack
18    Exit Sub

Err_Handler:
19    GlobalErrHandler
20    Resume Proc_Exit
End Sub

Private Sub lstWholesaler_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|lstWholesaler_Change"

3     Call ReCalcQA3ProductValues

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub cboVarPrice_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cboVarPrice_Change"

3     Call ReCalcQA3ProductValues

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub ReCalcQA3ProductValues()
      Dim i As Integer

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|ReCalcQA3ProductValues"

      ' Loop through the Products list to update values
3     With lstProducts
4         If lstProducts.ListCount <> 0 Then
5             For i = 0 To lstProducts.ListCount - 1
                  ' Contracted GSV
6                 .List(i, ProdList_ContractGSV) = CalcForecastGSV(.List(i, ProdList_ProdCode), cboVarPrice.Text, SetEmptyValue(.List(i, ProdList_ContractCases), ZeroValue))
7                 .List(i, ProdList_ContractGSVRoundoff) = RoundNum(.List(i, ProdList_ContractGSV), 0)
8                 lstTrdTerms.List(i, TTList_ContractGSV) = .List(i, ProdList_ContractGSVRoundoff)
9                 lstQA3.List(i, QA3List_ContractGSV) = .List(i, ProdList_ContractGSVRoundoff)

                  ' Direct Price
10                If cboRouteToMarket = "Direct" Then
11                    lstQA3.List(i, QA3List_DirectPrice) = GetVariablePricing(cboVarPrice.Text, .List(i, ProdList_ProdCode))
12                    lstQA3.List(i, QA3List_DirectPriceRoundoff) = RoundNum(lstQA3.List(i, QA3List_DirectPrice), 2)
13                End If

                  ' NIP Price or LUC Auto
14                lstQA3.List(i, QA3List_NipOrLUCAuto) = CalcNIPOrLUC(lstQA3.List(i, QA3List_Family), GetDirectOrWholeSalePrice(i), GetAdminAndFreight, lstQA3.List(i, QA3List_QA3Input), .List(i, ProdList_UnitsPerCase), .List(i, ProdList_BottleSize), lstTrdTerms.List(i, TTList_TTLtr), lstTrdTerms.List(i, TTList_TTGSV))
15                lstQA3.List(i, QA3List_NipOrLUCAutoRoundoff) = RoundNum(lstQA3.List(i, QA3List_NipOrLUCAuto), 2)
                  
                  ' QA3 Auto
16                lstQA3.List(i, QA3List_QA3Auto) = CalcQA3Auto(lstQA3.List(i, QA3List_Family), GetDirectOrWholeSalePrice(i), GetAdminAndFreight, lstQA3.List(i, QA3List_NipOrLUCInput), .List(i, ProdList_UnitsPerCase), .List(i, ProdList_BottleSize), lstTrdTerms.List(i, TTList_TTLtr), lstTrdTerms.List(i, TTList_TTGSV))
17                lstQA3.List(i, QA3List_QA3AutoRoundoff) = RoundNum(lstQA3.List(i, QA3List_QA3Auto), 2)
                  
                  ' KWI
18                lstQA3.List(i, QA3List_KWI) = GetKWI(SetEmptyValue(.List(i, ProdList_ContractGSV), ZeroValue), lstWholesaler, .List(i, ProdList_ProdCode))
19                lstQA3.List(i, QA3List_KWIRoundoff) = RoundNum(lstQA3.List(i, QA3List_KWI), 2)
                  
                  ' COP
20                lstQA3.List(i, QA3List_COP) = GetCOP(SetEmptyValue(.List(i, ProdList_ContractGSV), ZeroValue), SetEmptyValue(lstQA3.List(i, QA3List_QA3), ZeroValue), lstWholesaler, .List(i, ProdList_ProdCode))
21                lstQA3.List(i, QA3List_COPRoundoff) = RoundNum(lstQA3.List(i, QA3List_COP), 2)
22            Next i
23        End If
24    End With

Proc_Exit:
25    PopCallStack
26    Exit Sub

Err_Handler:
27    GlobalErrHandler
28    Resume Proc_Exit
End Sub

' QA3 Only Input ----------------------------------------------------------------------Start
Private Sub txtNIPOrLUCUser_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtNIPOrLUCUser_Change"

3     Call ToggleTextBoxEdit(txtQA3PerCaseUser, txtNIPOrLUCUser)

4     txtQA3PerCaseAuto = CalcQA3Auto(txtQA3Family, txtDirectPrice + txtWholesalerPrice, GetAdminAndFreight, txtNIPOrLUCUser, txtProdUnitsPerCase, txtProdBottleSize, txtTrdTermsPerLtr, txtTrdTermsGSV)

Proc_Exit:
5     PopCallStack
6     Exit Sub

Err_Handler:
7     GlobalErrHandler
8     Resume Proc_Exit
End Sub

Private Sub txtWholesalerPrice_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtWholesalerPrice_Change"

3     txtNIPOrLUCAuto = CalcNIPOrLUC(txtFamily, txtDirectPrice + txtWholesalerPrice, GetAdminAndFreight, txtQA3PerCaseUser, txtUnitsPerCase, txtBottleSize, txtTrdTermsPerLtr, txtTrdTermsGSV)
4     txtQA3PerCaseAuto = CalcQA3Auto(txtFamily, txtDirectPrice + txtWholesalerPrice, GetAdminAndFreight, txtNIPOrLUCUser, txtUnitsPerCase, txtBottleSize, txtTrdTermsPerLtr, txtTrdTermsGSV)

Proc_Exit:
5     PopCallStack
6     Exit Sub

Err_Handler:
7     GlobalErrHandler
8     Resume Proc_Exit
End Sub

Private Sub txtQA3PerCaseUser_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtQA3PerCaseUser_Change"

3     Call ToggleTextBoxEdit(txtQA3PerCaseUser, txtNIPOrLUCUser)

4     txtNIPOrLUCAuto = CalcNIPOrLUC(txtQA3Family, txtDirectPrice + txtWholesalerPrice, GetAdminAndFreight, txtQA3PerCaseUser, txtProdUnitsPerCase, txtProdBottleSize, txtTrdTermsPerLtr, txtTrdTermsGSV)

Proc_Exit:
5     PopCallStack
6     Exit Sub

Err_Handler:
7     GlobalErrHandler
8     Resume Proc_Exit
End Sub

Private Function CalcQA3Auto(strFamily As String, varVarPrice As Variant, varAdminAndFreight As Variant, _
                             varNIPOrLUC As Variant, varUnitsPerCase As Variant, varSize As Variant, _
                             varTTPerLtr As Variant, varTTGSV As Variant) As Variant

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|CalcQA3Auto"

3     CalcQA3Auto = 0
4     If IsNumeric(SetEmptyValue(varNIPOrLUC, NullStrings)) Then

      '    varTTPerLtr = SetEmptyValue(varTTPerLtr, ZeroValue)
      '    varTTGSV = SetEmptyValue(varTTGSV, ZeroValue)

5         Select Case UCase(strFamily)
              Case "SPIRITS"
                  'CalcQA3Auto = varVarPrice - ((((varNIPOrLUC * (varSize / g_dblNIP_Const)) * varUnitsPerCase) + varRebate) - varAdminAndFreight)
                  'CalcQA3Auto = varVarPrice - ((((varNIPOrLUC * (varSize / g_dblNIP_Const)) * varUnitsPerCase) - varAdminAndFreight - (varTTPerLtr / ((varSize / 1000) * varUnitsPerCase))) / (1 - (varTTGSV / 100)))
6                 CalcQA3Auto = varVarPrice - (((varNIPOrLUC * (varSize / g_dblNIP_Const)) * varUnitsPerCase) - varAdminAndFreight)
7             Case "WINE"
                  'CalcQA3Auto = varVarPrice - (((varNIPOrLUC * varUnitsPerCase) + varRebate) - varAdminAndFreight) / g_dblWET
                  'CalcQA3Auto = varVarPrice - ((varNIPOrLUC * varUnitsPerCase) + (varTTPerLtr / ((varSize / 1000) * varUnitsPerCase)) - varAdminAndFreight) / (g_dblWET - (varTTGSV / 100))
8                 CalcQA3Auto = varVarPrice - ((varNIPOrLUC * varUnitsPerCase) - varAdminAndFreight) / g_dblWET
9             Case "OTHER ALCOHOLIC BEVERAGES"
                  'CalcQA3Auto = varVarPrice - (((varNIPOrLUC * varUnitsPerCase) + varRebate) - varAdminAndFreight) / g_dblWET
                  'CalcQA3Auto = varVarPrice - ((varNIPOrLUC * varUnitsPerCase) + (varTTPerLtr / ((varSize / 1000) * varUnitsPerCase)) - varAdminAndFreight) / (g_dblWET - (varTTGSV / 100))
10                CalcQA3Auto = varVarPrice - ((varNIPOrLUC * varUnitsPerCase) - varAdminAndFreight)
11        End Select
12    End If

13    If CalcQA3Auto = 0 Then CalcQA3Auto = vbNullString

Proc_Exit:
14    PopCallStack
15    Exit Function

Err_Handler:
16    GlobalErrHandler
17    Resume Proc_Exit
End Function

Private Function CalcQA3(varQA3perCase As Variant, varLtr As Variant, varBottleSize As Variant, varUnitsPerCase As Variant)

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|CalcQA3"

3     CalcQA3 = 0
4     If IsNumeric(SetEmptyValue(varQA3perCase, NullStrings)) Then
5         CalcQA3 = SetEmptyValue(varQA3perCase, ZeroValue) * (varLtr / ((varBottleSize / 1000) * varUnitsPerCase))
6     End If

7     If CalcQA3 = 0 Then CalcQA3 = vbNullString

Proc_Exit:
8     PopCallStack
9     Exit Function

Err_Handler:
10    GlobalErrHandler
11    Resume Proc_Exit
End Function

Private Function CalcTradingTerms(varTT_L As Variant, varVol As Variant, varTT_Pct_GSV As Variant, varGSV As Variant, varQA3 As Variant) As Variant
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|CalcTradingTerms"

3     CalcTradingTerms = 0

4     If SetEmptyValue(varTT_L, ZeroValue) = 0 And SetEmptyValue(varTT_Pct_GSV, ZeroValue) = 0 Then
5         CalcTradingTerms = 0
6     Else
7         CalcTradingTerms = (SetEmptyValue(varTT_L, ZeroValue) * SetEmptyValue(varVol, ZeroValue)) + (SetEmptyValue(varTT_Pct_GSV, ZeroValue) / 100) * (SetEmptyValue(varGSV, ZeroValue) - SetEmptyValue(varQA3, ZeroValue))
8     End If

9     If CalcTradingTerms = 0 Then CalcTradingTerms = vbNullString

Proc_Exit:
10    PopCallStack
11    Exit Function

Err_Handler:
12    GlobalErrHandler
13    Resume Proc_Exit
End Function

'Private Function CalcTradingTerms(varTT_L As Variant, varBottleSize As Variant, varUnitsPerCase As Variant, varTT_Pct_GSV As Variant, varGSV As Variant, varQA3 As Variant) As Variant
'If gEnableErrorHandling Then On Error GoTo Err_Handler
'PushCallStack "frmMain|CalcTradingTerms"
'
'CalcTradingTerms = 0
'
'CalcTradingTerms = (SetEmptyValue(varTT_L, ZeroValue) / ((varBottleSize / 1000) * varUnitsPerCase)) + (SetEmptyValue(varTT_Pct_GSV, ZeroValue) / 100) * (SetEmptyValue(varGSV, ZeroValue) - SetEmptyValue(varQA3, ZeroValue))
'
'If CalcTradingTerms = 0 Then CalcTradingTerms = vbNullString
'
'Proc_Exit:
'PopCallStack
'Exit Function
'
'Err_Handler:
'GlobalErrHandler
'Resume Proc_Exit
'End Function

Private Function CalcAdditionalTerms(varTT_L As Variant, varMaxTT_L As Variant, varLtr As Variant, varTT_Pct_GSV As Variant, varMaxTT_Pct_GSV As Variant, varGSV As Variant, varQA3 As Variant) As Variant
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|CalcAdditionalTerms"

3     CalcAdditionalTerms = 0

4     If SetEmptyValue(varMaxTT_L, ZeroValue) = 0 And SetEmptyValue(varMaxTT_Pct_GSV, ZeroValue) = 0 Then
5         CalcAdditionalTerms = 0
6     Else
7         CalcAdditionalTerms = ((SetEmptyValue(varMaxTT_L, ZeroValue) - SetEmptyValue(varTT_L, ZeroValue)) * varLtr) + ((SetEmptyValue(varMaxTT_Pct_GSV, ZeroValue) / 100 - SetEmptyValue(varTT_Pct_GSV, ZeroValue) / 100) * (varGSV - SetEmptyValue(varQA3, ZeroValue)))
8     End If

9     If CalcAdditionalTerms = 0 Then CalcAdditionalTerms = vbNullString

Proc_Exit:
10    PopCallStack
11    Exit Function

Err_Handler:
12    GlobalErrHandler
13    Resume Proc_Exit
End Function

Private Function CalcBannerTerms(varTTBannerGSV As Variant, varTTBannerGSVlessQA3 As Variant, varGSV As Variant, varQA3 As Variant) As Variant
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|CalcBannerTerms"

3     CalcBannerTerms = 0

4     If SetEmptyValue(varTTBannerGSV, ZeroValue) = 0 And SetEmptyValue(varTTBannerGSVlessQA3, ZeroValue) = 0 Then
5         CalcBannerTerms = 0
6     Else
7         CalcBannerTerms = ((SetEmptyValue(varTTBannerGSV, ZeroValue) / 100) * varGSV) + ((SetEmptyValue(varTTBannerGSVlessQA3, ZeroValue) / 100) * (SetEmptyValue(varGSV, ZeroValue) - SetEmptyValue(varQA3, ZeroValue)))
8     End If

9     If CalcBannerTerms = 0 Then CalcBannerTerms = vbNullString

Proc_Exit:
10    PopCallStack
11    Exit Function

Err_Handler:
12    GlobalErrHandler
13    Resume Proc_Exit
End Function

Private Function CalcNIPOrLUC(strFamily As String, varVarPrice As Variant, varAdminAndFreight As Variant, _
                              varQA3perCase As Variant, varUnitsPerCase As Variant, varSize As Variant, _
                              varTTPerLtr As Variant, varTTGSV As Variant) As Variant
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|CalcNIPOrLUC"

3     CalcNIPOrLUC = 0

4     varTTPerLtr = SetEmptyValue(varTTPerLtr, ZeroValue)
5     varTTGSV = SetEmptyValue(varTTGSV, ZeroValue)

6     Select Case UCase(strFamily)
          Case "SPIRITS"
7             CalcNIPOrLUC = CalcNIPPrice(CalcLUC(varVarPrice, varQA3perCase, varSize, varUnitsPerCase, 1, varAdminAndFreight, varTTPerLtr, varTTGSV), varSize)
8         Case "WINE"
9             CalcNIPOrLUC = CalcLUC(varVarPrice, varQA3perCase, varSize, varUnitsPerCase, g_dblWET, varAdminAndFreight, varTTPerLtr, varTTGSV)
10        Case "OTHER ALCOHOLIC BEVERAGES"
11            CalcNIPOrLUC = CalcLUC(varVarPrice, varQA3perCase, varSize, varUnitsPerCase, 1, varAdminAndFreight, varTTPerLtr, varTTGSV)
12    End Select

13    If CalcNIPOrLUC = 0 Then CalcNIPOrLUC = vbNullString

Proc_Exit:
14    PopCallStack
15    Exit Function

Err_Handler:
16    GlobalErrHandler
17    Resume Proc_Exit
End Function
' QA3 Only Input ----------------------------------------------------------------------End

' COOP and A&P ------------------------------------------------------------------------Start
' Calculate COOP and A&P totals
Private Sub txtCoopCashPay_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtCoopCashPay_Change"

3     txtTotalCashPay = AddToTotalFund(txtCoopCashPay, txtAnPCashPay)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub txtAnPCashPay_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtAnPCashPay_Change"

3     txtTotalCashPay = AddToTotalFund(txtCoopCashPay, txtAnPCashPay)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub txtCoopBonusStock_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtCoopBonusStock_Change"

3     txtTotalBonusStock = AddToTotalFund(txtCoopBonusStock, txtAnPBonusStock)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub txtAnPBonusStock_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtAnPBonusStock_Change"

3     txtTotalBonusStock = AddToTotalFund(txtCoopBonusStock, txtAnPBonusStock)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub txtCoopPromoFund_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtCoopPromoFund_Change"

3     txtTotalPromoFund = AddToTotalFund(txtCoopPromoFund, txtAnPPromoFund)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub txtAnPPromoFund_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtAnPPromoFund_Change"

3     txtTotalPromoFund = AddToTotalFund(txtCoopPromoFund, txtAnPPromoFund)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub txtCoopStaffIncentives_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtCoopStaffIncentives_Change"

3     txtTotalStaffIncentives = AddToTotalFund(txtCoopStaffIncentives, txtAnPStaffIncentives)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub txtAnPStaffIncentives_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtAnPStaffIncentives_Change"

3     txtTotalStaffIncentives = AddToTotalFund(txtCoopStaffIncentives, txtAnPStaffIncentives)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub txtCoopPRAHospitality_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtCoopPRAHospitality_Change"

3     txtTotalPRAHospitality = AddToTotalFund(txtCoopPRAHospitality, txtAnPPRAHospitality)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub txtAnPPRAHospitality_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtAnPPRAHospitality_Change"

3     txtTotalPRAHospitality = AddToTotalFund(txtCoopPRAHospitality, txtAnPPRAHospitality)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Function AddToTotalFund(txtCoop As MSForms.TextBox, txtAnP As MSForms.TextBox) As Variant
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|AddToTotalFund"

3     AddToTotalFund = vbNullString

4     If HasInvalidNumber(txtCoop, True, "COOP amount") Then GoTo Validation_Error
5     If HasInvalidNumber(txtAnP, True, "A&P amount") Then GoTo Validation_Error

6     AddToTotalFund = CDbl(SetEmptyValue(txtCoop, ZeroValue)) + CDbl(SetEmptyValue(txtAnP, ZeroValue))

7     If AddToTotalFund = 0 Then AddToTotalFund = vbNullString

Proc_Exit:
8     PopCallStack
9     Exit Function

Validation_Error:
10    GoTo Proc_Exit

Err_Handler:
11    GlobalErrHandler
12    Resume Proc_Exit

End Function

Private Sub cmdClearCOOPAndAnP_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdClearCOOPAndAnP_Click"

3     txtCoopCashPay = vbNullString
4     txtCoopBonusStock = vbNullString
5     txtCoopPromoFund = vbNullString
6     txtCoopStaffIncentives = vbNullString
7     txtCoopPRAHospitality = vbNullString
8     txtAnPCashPay = vbNullString
9     txtAnPBonusStock = vbNullString
10    txtAnPPromoFund = vbNullString
11    txtAnPStaffIncentives = vbNullString
12    txtAnPPRAHospitality = vbNullString
13    txtTotalCashPay = vbNullString
14    txtTotalBonusStock = vbNullString
15    txtTotalPromoFund = vbNullString
16    txtTotalStaffIncentives = vbNullString
17    txtTotalPRAHospitality = vbNullString
18    txtReciprocalSpend = vbNullString
19    txtCommentsCashPay = vbNullString
20    txtCommentsBonusStock = vbNullString
21    txtCommentsPromoFund = vbNullString
22    txtCommentsStaffIncentives = vbNullString
23    txtCommentsPRAHospitality = vbNullString
24    txtReciprocalSpendComments = vbNullString

Proc_Exit:
25    PopCallStack
26    Exit Sub

Err_Handler:
27    GlobalErrHandler
28    Resume Proc_Exit
End Sub

Private Sub txtReciprocalSpend_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtReciprocalSpend_Change"

3     If HasInvalidNumber(txtReciprocalSpend, True, "Reciprocal Spend") Then GoTo Validation_Error

Proc_Exit:
4     PopCallStack
5     Exit Sub

Validation_Error:
6     GoTo Proc_Exit

Err_Handler:
7     GlobalErrHandler
8     Resume Proc_Exit
End Sub

' COOP and A&P ------------------------------------------------------------------------End

' Trading Terms Options ---------------------------------------------------------------Start
Private Sub optContractAndNonContract_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|optContractAndNonContract_Click"

3     Call EnableManualTradingTermsEdit(False)

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit
End Sub

Private Sub optEnterManually_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|optEnterManually_Click"

3     txtAllProd_PctGSV = vbNullString
4     txtAllProd_DollarPerLitre = vbNullString
5     cboAllProd_FreqOfPayments = vbNullString
6     Call EnableManualTradingTermsEdit(True)

Proc_Exit:
7     PopCallStack
8     Exit Sub

Err_Handler:
9     GlobalErrHandler
10    Resume Proc_Exit
End Sub

Private Sub EnableManualTradingTermsEdit(blnEnabled As Boolean)

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|EnableManualTradingTermsEdit"

3     cmdEditTradingTerms.Enabled = blnEnabled

4     txtAllProd_PctGSV.Enabled = Not blnEnabled
5     txtAllProd_DollarPerLitre.Enabled = Not blnEnabled
6     cboAllProd_FreqOfPayments.Enabled = Not blnEnabled
7     cmdApplyToAllProducts.Enabled = Not blnEnabled

Proc_Exit:
8     PopCallStack
9     Exit Sub

Err_Handler:
10    GlobalErrHandler
11    Resume Proc_Exit
End Sub
' Trading Terms Options ---------------------------------------------------------------End

' Edit Trading Terms-------------------------------------------------------------------Start
Private Sub cmdApplyToAllProducts_Click()
      Dim i As Long

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdApplyToAllProducts_Click"

      ' Validate inputs
3     If HasInvalidNumber(txtAllProd_DollarPerLitre, True, strName:="$ Per Litre") Then GoTo Proc_Exit
4     If HasInvalidNumber(txtAllProd_PctGSV, True, strName:="% GSV-QA3") Then GoTo Proc_Exit
5     If HasInvalidNumber(txtAllProd_DollarPerLitre) = False Or HasInvalidNumber(txtAllProd_PctGSV) = False Then
6         If HasEmptyValues(cboAllProd_FreqOfPayments, strName:="Frequency of Payments") Then GoTo Proc_Exit
7     End If

8     If Len(txtAllProd_DollarPerLitre) = 0 And Len(txtAllProd_PctGSV) = 0 And Len(txtAllProd_DollarPerLitre) = 0 Then
9         If MsgBox("This will delete all Standard Terms values.", vbOKCancel) <> vbOK Then
10            GoTo Proc_Exit
11        End If
12    End If

      ' Apply inputs and Recalculate Terms
13    With lstTrdTerms
14        If .ListCount > 0 Then
15            For i = 0 To .ListCount - 1
16                .List(i, TTList_TTLtr) = txtAllProd_DollarPerLitre
17                .List(i, TTList_TTGSV) = txtAllProd_PctGSV
18                .List(i, TTList_FreqOfPayment) = cboAllProd_FreqOfPayments
                  
                  ' Recalc Terms
19                .List(i, TTList_StandardTerm) = CalcTradingTerms(.List(i, TTList_TTLtr), lstProducts.List(i, ProdList_ContractVol), .List(i, TTList_TTGSV), lstProducts.List(i, ProdList_ContractGSV), lstQA3.List(i, QA3List_QA3))
20                .List(i, TTList_AddnlTerm) = CalcAdditionalTerms(.List(i, TTList_TTLtr), .List(i, TTList_TTMaxLtr), lstProducts.List(i, ProdList_ContractVol), .List(i, TTList_TTGSV), .List(i, TTList_TTMaxGSV), lstProducts.List(i, ProdList_ContractGSV), lstQA3.List(i, QA3List_QA3))
21            Next i
22        End If
23    End With

Proc_Exit:
24    PopCallStack
25    Exit Sub

Err_Handler:
26    GlobalErrHandler
27    Resume Proc_Exit
End Sub

Private Sub cmdApplyToCondTerms_Click()
      Dim i As Long

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdApplyToCondTerms_Click"

      ' Validate inputs
3     If HasInvalidNumber(txtAllCondMaxLitre, True, strName:="MAX$ Per Litre") Then GoTo Proc_Exit
4     If HasInvalidNumber(txtAllCondMaxGSV, True, strName:="MAX% GSV-QA3") Then GoTo Proc_Exit
5     If HasInvalidNumber(txtAllCondMaxLitre) = False Or HasInvalidNumber(txtAllCondMaxGSV) = False Then
6         If HasEmptyValues(txtAllCondComment, strName:="Conditional Terms Comments") Then GoTo Proc_Exit
7     End If

8     If Len(txtAllCondMaxLitre) = 0 And Len(txtAllCondMaxGSV) = 0 And Len(txtAllCondComment) = 0 Then
9         If MsgBox("This will delete all Conditional Terms values.", vbOKCancel) <> vbOK Then
10            GoTo Proc_Exit
11        End If
12    End If

      ' Apply inputs and recalc Terms
13    With lstTrdTerms
14        If .ListCount > 0 Then
15            For i = 0 To .ListCount - 1
16                .List(i, TTList_TTMaxLtr) = txtAllCondMaxLitre
17                .List(i, TTList_TTMaxGSV) = txtAllCondMaxGSV
18                .List(i, TTList_TTCondComment) = txtAllCondComment
                  
                  ' Recalc Terms
19                .List(i, TTList_StandardTerm) = CalcTradingTerms(.List(i, TTList_TTLtr), lstProducts.List(i, ProdList_ContractVol), .List(i, TTList_TTGSV), lstProducts.List(i, ProdList_ContractGSV), lstQA3.List(i, QA3List_QA3))
20                .List(i, TTList_AddnlTerm) = CalcAdditionalTerms(.List(i, TTList_TTLtr), .List(i, TTList_TTMaxLtr), lstProducts.List(i, ProdList_ContractVol), .List(i, TTList_TTGSV), .List(i, TTList_TTMaxGSV), lstProducts.List(i, ProdList_ContractGSV), lstQA3.List(i, QA3List_QA3))
21            Next i
22        End If
23    End With

Proc_Exit:
24    PopCallStack
25    Exit Sub

Err_Handler:
26    GlobalErrHandler
27    Resume Proc_Exit
End Sub

Private Sub cmdEditTradingTerms_Click()
      Dim idx As Integer

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdEditTradingTerms_Click"

      ' Check if there are record selected
3     If LBXSelectCount(lstTrdTerms) = 1 Then
4         Select Case cmdEditTradingTerms.Caption
              Case "Edit"
                  ' Change button caption
5                 cmdEditTradingTerms.Caption = "Update"
                  
                  ' Clear inputs
6                 Call cmdClearInputsTradingTerms_Click
                  
                  ' Disable other controls
7                 cmdClearInputsTradingTerms.Enabled = True
8                 fraRecordCommands.Enabled = False
9                 fraAllStanTerms.Enabled = False
10                fraAllCondTerms.Enabled = False
11                lstTrdTerms.Locked = True
12                mpgeOP.Pages(0).Enabled = False
13                mpgeOP.Pages(1).Enabled = False
14                mpgeOP.Pages(3).Enabled = False
15                mpgeOP.Pages(4).Enabled = False
16                mpgeOP.Pages(5).Enabled = False
                  
                  ' Populate fields
17                With lstTrdTerms
18                    txtTTProdType = .List(.ListIndex, TTList_ProdType)
19                    txtTTProdBrand = .List(.ListIndex, TTList_Brand)
20                    txtTTProdDesc = .List(.ListIndex, TTList_ProdDesc)
21                    txtTTLitres = .List(.ListIndex, TTList_ContractVol)
22                    txtTTGSV = .List(.ListIndex, TTList_ContractGSV)
23                    txtDollarPerLitre = .List(.ListIndex, TTList_TTLtr)
24                    txtTTPctOfGSV = .List(.ListIndex, TTList_TTGSV)
25                    cboTTFreqOfPayments = .List(.ListIndex, TTList_FreqOfPayment)
26                    txtTTAddnlDollarPerLitre = .List(.ListIndex, TTList_TTMaxLtr)
27                    txtTTAddnlPctOfGSV = .List(.ListIndex, TTList_TTMaxGSV)
28                    txtTTCondComments = .List(.ListIndex, TTList_TTCondComment)
29                End With
                  
30            Case "Update"
                  ' Validate inputs
31                If HasInvalidNumber(txtDollarPerLitre, True, strName:="Trading Terms: $ Per Litre") Then GoTo Proc_Exit
32                If HasInvalidNumber(txtTTPctOfGSV, True, strName:="Trading Terms: % of GSV-QA3") Then GoTo Proc_Exit
33                If HasEmptyValues(cboTTFreqOfPayments, strName:="Frequency of Payments") Then GoTo Proc_Exit
34                If HasInvalidNumber(txtTTAddnlDollarPerLitre, True, strName:="Trading Terms: Additional $ per Litre above Contracted Litres") Then GoTo Proc_Exit
35                If HasInvalidNumber(txtTTAddnlPctOfGSV, True, strName:="Trading Terms: Additional % of GSV-QA3 above Contracted GSV") Then GoTo Proc_Exit
                  
                  ' Get index of selected item
36                idx = lstTrdTerms.ListIndex
                  
                  ' Update listbox values
37                With lstTrdTerms
38                    .List(idx, TTList_TTLtr) = txtDollarPerLitre
39                    .List(idx, TTList_TTGSV) = txtTTPctOfGSV
40                    .List(idx, TTList_FreqOfPayment) = cboTTFreqOfPayments
41                    .List(idx, TTList_TTMaxLtr) = txtTTAddnlDollarPerLitre
42                    .List(idx, TTList_TTMaxGSV) = txtTTAddnlPctOfGSV
43                    .List(idx, TTList_TTCondComment) = txtTTCondComments
44                    .List(idx, TTList_StandardTerm) = CalcTradingTerms(txtDollarPerLitre, lstProducts.List(idx, ProdList_ContractVol), txtTTPctOfGSV, lstProducts.List(idx, ProdList_ContractGSV), lstQA3.List(idx, QA3List_QA3))
45                    .List(idx, TTList_AddnlTerm) = CalcAdditionalTerms(txtDollarPerLitre, txtTTAddnlDollarPerLitre, lstProducts.List(idx, ProdList_ContractVol), txtTTPctOfGSV, txtTTAddnlPctOfGSV, lstProducts.List(idx, ProdList_ContractGSV), lstQA3.List(idx, QA3List_QA3))
46                    .List(idx, TTList_BannerTerm) = CalcBannerTerms(txtTTBannerGSV, txtTTBannerGSVlessQA3, lstProducts.List(idx, ProdList_ContractGSV), lstQA3.List(idx, QA3List_QA3))
47                End With
                  
48                With lstQA3
                      ' Recalculate NIP Price or LUC Auto
49                    .List(idx, QA3List_NipOrLUCAuto) = CalcNIPOrLUC(lstQA3.List(idx, QA3List_Family), GetDirectOrWholeSalePrice(idx), GetAdminAndFreight, .List(idx, QA3List_QA3Input), lstProducts.List(idx, ProdList_UnitsPerCase), lstProducts.List(idx, ProdList_BottleSize), txtDollarPerLitre, txtTTPctOfGSV)
50                    .List(idx, QA3List_NipOrLUCAutoRoundoff) = RoundNum(.List(idx, QA3List_NipOrLUCAuto), 2)
                  
                      ' Recalculate QA3 Auto
51                    .List(idx, QA3List_QA3Auto) = CalcQA3Auto(lstQA3.List(idx, QA3List_Family), GetDirectOrWholeSalePrice(idx), GetAdminAndFreight, .List(idx, QA3List_NipOrLUCInput), lstProducts.List(idx, ProdList_UnitsPerCase), lstProducts.List(idx, ProdList_BottleSize), txtDollarPerLitre, txtTTPctOfGSV)
52                    .List(idx, QA3List_QA3AutoRoundoff) = RoundNum(.List(idx, QA3List_QA3Auto), 2)
53                End With
                
                  ' Change button caption
54                cmdEditTradingTerms.Caption = "Edit"
                  
                  ' Clear inputs
55                Call cmdClearInputsTradingTerms_Click
                  
                  ' Enable back the controls
56                cmdClearInputsTradingTerms.Enabled = False
57                fraRecordCommands.Enabled = True
58                fraAllStanTerms.Enabled = True
59                fraAllCondTerms.Enabled = True
60                lstTrdTerms.Locked = False
61                mpgeOP.Pages(0).Enabled = True
62                mpgeOP.Pages(1).Enabled = True
63                mpgeOP.Pages(3).Enabled = True
64                mpgeOP.Pages(4).Enabled = True
65                mpgeOP.Pages(5).Enabled = True
66        End Select
67    End If

Proc_Exit:
68    PopCallStack
69    Exit Sub

Err_Handler:
70    GlobalErrHandler
71    Resume Proc_Exit
End Sub

Private Function GetAdminAndFreight() As Variant
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|GetAdminAndFreight"

      ' Add Admin and Freight if Wholesaler is ALM
3     GetAdminAndFreight = 0
4     If cboRouteToMarket = "Indirect" And LBXSelectCount(lstWholesaler) <> 0 Then
5         If InStr(1, GetIN_List(LBXSelectedItems(lstWholesaler, 0), vbNullString, "|"), "ALM") <> 0 Then
6             GetAdminAndFreight = g_dblALM_Admin + g_dblALM_Freight
7         End If
8     End If

Proc_Exit:
9     PopCallStack
10    Exit Function

Err_Handler:
11    GlobalErrHandler
12    Resume Proc_Exit
End Function

Private Function GetDirectOrWholeSalePrice(idx As Integer) As Variant
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|GetDirectOrWholeSalePrice"

3     GetDirectOrWholeSalePrice = 0

4     If lstQA3.ListCount <> 0 Then
5         Select Case cboRouteToMarket
              Case "Direct"
6                 GetDirectOrWholeSalePrice = SetEmptyValue(lstQA3.List(idx, QA3List_DirectPrice), ZeroValue)
7             Case "Indirect"
8                 GetDirectOrWholeSalePrice = SetEmptyValue(lstQA3.List(idx, QA3List_WSPrice), ZeroValue)
9         End Select
10    End If

Proc_Exit:
11    PopCallStack
12    Exit Function

Err_Handler:
13    GlobalErrHandler
14    Resume Proc_Exit
End Function

Private Function GetQA3FromList(idx As Integer) As Variant
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|GetQA3FromList"

3     GetQA3FromList = 0
4     With lstProducts
5         If SetEmptyValue(.List(idx, QA3List_QA3Input), ZeroValue) = 0 Then
6             GetQA3FromList = .List(idx, QA3List_QA3Auto)
7         ElseIf SetEmptyValue(.List(idx, QA3List_QA3Auto), ZeroValue) = 0 Then
8             GetQA3FromList = .List(idx, QA3List_QA3Input)
9         End If
10    End With

Proc_Exit:
11    PopCallStack
12    Exit Function

Err_Handler:
13    GlobalErrHandler
14    Resume Proc_Exit
End Function

Private Sub cmdClearInputsTradingTerms_Click()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdClearInputsTradingTerms_Click"

3     If cmdEditTradingTerms.Caption <> "Update" Then
4         txtTTProdType = vbNullString
5         txtTTProdBrand = vbNullString
6         txtTTProdDesc = vbNullString
7         txtTTLitres = vbNullString
8         txtTTGSV = vbNullString
9     End If

10    txtDollarPerLitre = vbNullString
11    txtTTPctOfGSV = vbNullString
12    cboTTFreqOfPayments = vbNullString
13    txtTTAddnlDollarPerLitre = vbNullString
14    txtTTAddnlPctOfGSV = vbNullString
15    txtTTCondComments = vbNullString

Proc_Exit:
16    PopCallStack
17    Exit Sub

Err_Handler:
18    GlobalErrHandler
19    Resume Proc_Exit
End Sub
' Edit Trading Terms-------------------------------------------------------------------End

' Trading Terms Textbox Edit Toggle----------------------------------------------------Start
Private Sub txtAllNonContract_PctGSV_Change()
1     Call ToggleTextBoxEdit(txtAllNonContract_PctGSV, txtAllNonContract_DollarPerLitre)
End Sub

Private Sub txtAllNonContract_DollarPerLitre_Change()
1     Call ToggleTextBoxEdit(txtAllNonContract_PctGSV, txtAllNonContract_DollarPerLitre)
End Sub

Private Sub txtAllProd_PctGSV_Change()
1     Call ToggleTextBoxEdit(txtAllProd_PctGSV, txtAllProd_DollarPerLitre, txtAllCondMaxGSV, txtAllCondMaxLitre)
End Sub

Private Sub txtAllProd_DollarPerLitre_Change()
1     Call ToggleTextBoxEdit(txtAllProd_PctGSV, txtAllProd_DollarPerLitre, txtAllCondMaxGSV, txtAllCondMaxLitre)
End Sub

Private Sub txtDollarPerLitre_Change()
1     Call ToggleTextBoxEdit(txtDollarPerLitre, txtTTPctOfGSV)
End Sub

Private Sub txtTTPctOfGSV_Change()
1     Call ToggleTextBoxEdit(txtDollarPerLitre, txtTTPctOfGSV)
End Sub

Private Sub txtTTAddnlDollarPerLitre_Change()
1     Call ToggleTextBoxEdit(txtTTAddnlDollarPerLitre, txtTTAddnlPctOfGSV)
End Sub

Private Sub txtTTAddnlPctOfGSV_Change()
1     Call ToggleTextBoxEdit(txtTTAddnlDollarPerLitre, txtTTAddnlPctOfGSV)
End Sub

Private Sub txtAllCondMaxGSV_Change()
1     Call ToggleTextBoxEdit(txtAllCondMaxGSV, txtAllCondMaxLitre, txtAllProd_PctGSV, txtAllProd_DollarPerLitre)
End Sub

Private Sub txtAllCondMaxLitre_Change()
1     Call ToggleTextBoxEdit(txtAllCondMaxGSV, txtAllCondMaxLitre, txtAllProd_PctGSV, txtAllProd_DollarPerLitre)
End Sub

Private Sub txtTTBannerGSV_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtTTBannerGSV_Change"

3     Call ToggleTextBoxEdit(txtTTBannerGSV, txtTTBannerGSVlessQA3)

4     If HasInvalidNumber(txtTTBannerGSV, True, "Banner Terms % GSV") Then GoTo Validation_Error

5     Call ReCalcBannerTerms(txtTTBannerGSV, txtTTBannerGSVlessQA3)

Proc_Exit:
6     PopCallStack
7     Exit Sub

Validation_Error:
8     GoTo Proc_Exit

Err_Handler:
9     GlobalErrHandler
10    Resume Proc_Exit
End Sub

Private Sub txtTTBannerGSVlessQA3_Change()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|txtTTBannerGSVlessQA3_Change"

3     Call ToggleTextBoxEdit(txtTTBannerGSV, txtTTBannerGSVlessQA3)

4     If HasInvalidNumber(txtTTBannerGSVlessQA3, True, "Banner Terms % GSV-QA3") Then GoTo Validation_Error

5     Call ReCalcBannerTerms(txtTTBannerGSV, txtTTBannerGSVlessQA3)

Proc_Exit:
6     PopCallStack
7     Exit Sub

Validation_Error:
8     GoTo Proc_Exit

Err_Handler:
9     GlobalErrHandler
10    Resume Proc_Exit
End Sub
' Trading Terms Textbox Edit Toggle----------------------------------------------------End

Private Sub ReCalcBannerTerms(varGSV As Variant, varGSVlessQA3 As Variant)
      Dim i As Integer

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|ReCalcBannerterms"

3     With lstTrdTerms
4         If .ListCount <> 0 Then
5             For i = 0 To .ListCount - 1
6                 .List(i, TTList_BannerTerm) = CalcBannerTerms(varGSV, varGSVlessQA3, lstProducts.List(i, ProdList_ContractGSV), lstQA3.List(i, QA3List_QA3))
7             Next i
8         End If
9     End With

Proc_Exit:
10    PopCallStack
11    Exit Sub

Err_Handler:
12    GlobalErrHandler
13    Resume Proc_Exit

End Sub

Private Sub cmdApprove_Click()
      Dim qry As String
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|cmdApprove_Click"

3     If MsgBox("Are you sure you want to approve the On Premise contract " & txtRefNumber & "?", vbYesNo, "Confirm Approval") = vbYes Then
          ' Update Status to Approved
4         qry = "UPDATE " & OP_MAIN_TBL & " " & _
                "SET StatusID = " & enumStatus.statApproved & " " & _
                "WHERE RefNumber = '" & txtRefNumber & "'"
5         cn.Execute qry
          
          ' Hide Contract details
6         Call HideDetails
7     End If

Proc_Exit:
8     PopCallStack
9     Exit Sub

Err_Handler:
10    GlobalErrHandler
11    Resume Proc_Exit
End Sub

Private Sub chkNonContract_Click()

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|chkNonContract_Click"

3     Select Case chkNonContract.Value
          Case -1
4             txtAllNonContract_PctGSV.Enabled = True
5             txtAllNonContract_DollarPerLitre.Enabled = True
6         Case 0
7             txtAllNonContract_PctGSV.Enabled = False
8             txtAllNonContract_DollarPerLitre.Enabled = False
9     End Select

Proc_Exit:
10    PopCallStack
11    Exit Sub

Err_Handler:
12    GlobalErrHandler
13    Resume Proc_Exit
End Sub

Private Sub chkBannerTerms_Click()

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "frmMain|chkBannerTerms_Click"

3     Select Case chkBannerTerms.Value
          Case -1
4             txtTTBannerGSV.Enabled = True
5             txtTTBannerGSVlessQA3.Enabled = True
6         Case 0
7             txtTTBannerGSV.Enabled = False
8             txtTTBannerGSVlessQA3.Enabled = False
9     End Select

Proc_Exit:
10    PopCallStack
11    Exit Sub

Err_Handler:
12    GlobalErrHandler
13    Resume Proc_Exit
End Sub


