Attribute VB_Name = "mGlobal"
Option Explicit

Public Const gEnableErrorHandling As Boolean = False

Public g_frmMain As frmMain

Global g_sLoginID As String
Global g_iAccessType As Integer
Global g_sDate As String
Global g_bForm As Boolean
Global strExitState As String   'Used to pass the reason for closing the promo form to the parent
Global g_blnUpdateInProgress As Boolean

Global g_dblNIP_Const As Double
Global g_dblWET As Double
Global g_dblALM_Admin As Double
Global g_dblALM_Freight As Double

Public Const TRANS_TBL = "Transactions"

Public Const SETTINGS_TBL = "T_Settings"

Public Const CUSTOMER_MAP_TBL = "T_Map_Customers"
Public Const OUTLET_MAP_TBL = "T_Map_Outlet"            ' Is it still used?
Public Const WHOLESALER_MAP_TBL = "T_Map_Wholesaler"
Public Const OUTLET_INFO_TBL = "T_Main_Outlet_Info"

Public Const PRODUCT_MAP_TBL = "T_Map_Products"
Public Const PRICING_MAP_TBL = "T_Map_Pricing"
Public Const EXCISE_MAP_TBL = "T_Map_Excise"
Public Const COGSPERLTR_MAP_TBL = "T_Map_COGSperLitre"
Public Const KWI_MAP_TBL = "T_Map_KWI"
Public Const COP_TERMS_MAP_TBL = "T_Map_COP_Terms"
Public Const ADJNAME_PREFIX_MAP_TBL = "T_Map_AdjName_Prefix"
Public Const MATCHCODE_ALMCUSTCODE_MAP_TBL = "T_Map_MatchCode_To_ALMCustCode"

Public Const PRA_EMPLOYEE_TBL = "T_PRA_Members"
Public Const PRA_MANAGER_TBL = "T_PRA_Managers"
Public Const STATUS_TBL = "T_MAP_Status"

' Temp sheets
Public Const PROG_DETAILS_TEMP_SHEET = "Programme_Details_Temp"
Public Const NON_QA3_INPUT_TEMP_SHEET = "Non_QA3_Input_Temp"
Public Const PEM_PREVIW_TEMP_SHEET = "PEM_Preview_Temp"
Public Const PEM_TEMP_SHEET = "PEM_Temp"
Public Const PEM_SUMM_TEMP_SHEET = "PEM_Summary_Temp"
Public Const PEM_TEMP_SHEET_RENAME = "Appendix Sheet"
Public Const PEM_SUMM_TEMP_SHEET_RENAME = "Summary Sheet"
Public Const E1_UPLOAD_TEMP_SHEET = "E1Upload_Temp"
Public Const E1_UPLOAD_TEMP_SHEET_RENAME = "E1 Upload"
Public Const DATA_DUMP_SHEET = "Data_Dump_Temp"
Public Const DATA_DUMP_SHEET_RENAME = "Data Dump"
Public Const ALM_DEAL_TEMP_SHEET = "ALM_Deal_Sheet_Temp"
Public Const ALM_DEAL_TEMP_SHEET_RENAME = "ALM Deal Sheet"
Public Const STANDARD_DEAL_TEMP_SHEET = "Standard_Deal_Sheet_Temp"

' Monthly Report templates
Public Const MONTHLY_NAT_CUST = "National_Customer_Temp"
Public Const MONTHLY_NAT_BRND = "National_Brand_Temp"
Public Const MONTHLY_MNGR_CUST_PERF = "Acct_Mgr_Cust_Perf_Temp"
Public Const MONTHLY_MNGR_CUST_PROD_PERF = "Acct_Mgr_Cust_Prod_Perf_Temp"

' Products tab listbox columns
Public Const ProdList_ProdType = 0
Public Const ProdList_Brand = 1
Public Const ProdList_BrandCode = 2
Public Const ProdList_Subbrand = 3
Public Const ProdList_SubbrandCode = 4
Public Const ProdList_ProdDesc = 5
Public Const ProdList_ProdCode = 6
Public Const ProdList_BottleSize = 7
Public Const ProdList_UnitsPerCase = 8
Public Const ProdList_ContractCases = 9
Public Const ProdList_ContractVol = 10
Public Const ProdList_ContractVolRoundoff = 11
Public Const ProdList_ContractGSV = 12
Public Const ProdList_ContractGSVRoundoff = 13
Public Const ProdList_Family = 14

' QA3 tab listbox columns
Public Const QA3List_ProdType = 0
Public Const QA3List_Brand = 1
Public Const QA3List_ProdDesc = 2
Public Const QA3List_ProdCode = 3
Public Const QA3List_ContractVol = 4
Public Const QA3List_ContractGSV = 5
Public Const QA3List_DirectPrice = 6
Public Const QA3List_DirectPriceRoundoff = 7
Public Const QA3List_WSPrice = 8
Public Const QA3List_WSPriceRoundoff = 9
Public Const QA3List_QA3Input = 10
Public Const QA3List_QA3InputRoundoff = 11
Public Const QA3List_NipOrLUCAuto = 12
Public Const QA3List_NipOrLUCAutoRoundoff = 13
Public Const QA3List_NipOrLUCInput = 14
Public Const QA3List_NipOrLUCInputRoundoff = 15
Public Const QA3List_QA3Auto = 16
Public Const QA3List_QA3AutoRoundoff = 17
Public Const QA3List_QA3 = 18
Public Const QA3List_QA3Roundoff = 19
Public Const QA3List_KWI = 20
Public Const QA3List_KWIRoundoff = 21
Public Const QA3List_COP = 22
Public Const QA3List_COPRoundoff = 23
Public Const QA3List_Family = 24

' Trading Terms tab listbox columns
Public Const TTList_ProdType = 0
Public Const TTList_Brand = 1
Public Const TTList_ProdDesc = 2
Public Const TTList_ProdCode = 3
Public Const TTList_ContractVol = 4
Public Const TTList_ContractGSV = 5
Public Const TTList_TTLtr = 6
Public Const TTList_TTGSV = 7
Public Const TTList_FreqOfPayment = 8
Public Const TTList_TTMaxLtr = 9
Public Const TTList_TTMaxGSV = 10
Public Const TTList_TTCondComment = 11
Public Const TTList_StandardTerm = 12
Public Const TTList_AddnlTerm = 13
Public Const TTList_BannerTerm = 14

Public Const DISABLE_COLOR = &HE0E0E0
Public Const ENABLE_COLOR = &H80000005

Public Const MAIN_SHEET = "Main"

Public wkb As Workbook

Public Enum enumIndexColor
    LightYellow = 6750207
    Yellow = 65535
    LightGreen = 5296274
    Green = 5287936
    White = 1
    LightBlue = 16764057
    Blue = 13408512
    LightGrey = 14540253
    Grey = 12632256
End Enum

Public Enum enumPromoDate
    Start_Date = 1
    End_Date = 2
End Enum

Public Enum enumStatus
    statDraft = 1
    statForApproval = 2
    statApproved = 3
    statView = 4
    statDeleted = 5
End Enum

Public Enum enumUserPermission  ' same as [T_User_Permission]
    OrdinaryUser = 1
    Admin = 2
    Manager = 3
End Enum

Public Sub LoadMain()
    
    ' Initialize
    Call StartupRoutine
    
    Set g_frmMain = New frmMain
    Load g_frmMain
    g_frmMain.Show
    
    'Application.OnKey "^{+}", "ToggleShowExcelApp"
    'Call ToggleShowExcelApp
End Sub

Public Sub ToggleShowExcelApp()
    If Application.Visible = False Then
        Application.Visible = True
    Else
        Application.Visible = False
    End If
End Sub

Public Function GenerateRefNum() As String
Dim strUserID As String
Dim strRefNum As String

If gEnableErrorHandling Then On Error GoTo Err_Handler
PushCallStack "mGlobal|GenerateRefNum"

strUserID = GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "ID", "WinLoginName", UCase(Environ("UserName")), """")
Do
    Randomize
    strRefNum = strUserID & "-" & Format(Now, "yymm") & "-" & Int((999 - 100 + 1) * Rnd + 100)
Loop Until Len(GetItemFromMappingTbl(OP_MAIN_TBL, "RefNumber", "RefNumber", strRefNum, """")) = 0

GenerateRefNum = strRefNum

Proc_Exit:
PopCallStack
Exit Function

Err_Handler:
GlobalErrHandler
Resume Proc_Exit
End Function

Public Sub UnloadLoadMain()
    Unload g_frmMain
End Sub

Public Sub StartupRoutine()
    ' Do something
End Sub

