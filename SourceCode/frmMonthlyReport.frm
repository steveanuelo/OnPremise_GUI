VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMonthlyReport 
   Caption         =   "Monthly Report"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5400
   OleObjectBlob   =   "frmMonthlyReport.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMonthlyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
If gEnableErrorHandling Then On Error GoTo Err_Handler
PushCallStack "frmMonthlyReport|UserForm_Activate"

Call UserForm_Initialize

Proc_Exit:
PopCallStack
Exit Sub

Err_Handler:
GlobalErrHandler
Resume Proc_Exit
End Sub

Private Sub UserForm_Initialize()
If gEnableErrorHandling Then On Error GoTo Err_Handler
PushCallStack "frmMonthlyReport|UserForm_Initialize"

' Month end selection
Call SetMonthEndSelection
'cboMonthEnd.Value = Format(CDate(WorksheetFunction.EoMonth(Now, 0)), "dd-mmm-yy")

' Report Type
With cboMontlyReportType
    .Clear
    .AddItem
    .List(.ListCount - 1, 0) = enumMonthlyReportType.NationalCustomer
    .List(.ListCount - 1, 1) = "National Customer"
    .AddItem
    .List(.ListCount - 1, 0) = enumMonthlyReportType.NationalBrand
    .List(.ListCount - 1, 1) = "National Brand"
    .AddItem
    .List(.ListCount - 1, 0) = enumMonthlyReportType.AccountManagerCustPerf
    .List(.ListCount - 1, 1) = "Account Manager-Customer Performance"
    .AddItem
    .List(.ListCount - 1, 0) = enumMonthlyReportType.AccountManagerCustAndProdPerf
    .List(.ListCount - 1, 1) = "Account Manager-Customer & Product Performance"
End With

cboMonthlyRptCreator.List = GetArrayList("SELECT DISTINCT T1.CreatorID, T2.Name FROM " & OP_MAIN_TBL & " AS T1 INNER JOIN " & PRA_EMPLOYEE_TBL & " AS T2 ON T1.CreatorID = T2.ID;", True)

cboMonthlyRptOutletGrp.List = GetArrayList("SELECT DISTINCT ContractLevelCode, OutletOrGroupName FROM " & OP_MAIN_TBL & ";", True)

cboStartEndDate.List = GetArrayList("SELECT FromDate, ToDate FROM " & OP_MAIN_TBL & ";", True)

lblMonthlyRptCreator.Enabled = False
cboMonthlyRptCreator.Enabled = False
lblMonthlyRptOutletGrp.Enabled = False
cboMonthlyRptOutletGrp.Enabled = False
lblStartEndDate.Enabled = False
cboStartEndDate.Enabled = False

Proc_Exit:
PopCallStack
Exit Sub

Err_Handler:
GlobalErrHandler
Resume Proc_Exit

End Sub

Private Sub cboMontlyReportType_Change()
If gEnableErrorHandling Then On Error GoTo Err_Handler
PushCallStack "frmMonthlyReport|cboMontlyReportType_Change"

Select Case cboMontlyReportType.Value
    Case enumMonthlyReportType.NationalCustomer
        lblMonthlyRptCreator.Enabled = False
        cboMonthlyRptCreator.Enabled = False
        lblMonthlyRptOutletGrp.Enabled = False
        cboMonthlyRptOutletGrp.Enabled = False
        lblStartEndDate.Enabled = False
        cboStartEndDate.Enabled = False
        
    Case enumMonthlyReportType.NationalBrand
        lblMonthlyRptCreator.Enabled = False
        cboMonthlyRptCreator.Enabled = False
        lblMonthlyRptOutletGrp.Enabled = False
        cboMonthlyRptOutletGrp.Enabled = False
        lblStartEndDate.Enabled = False
        cboStartEndDate.Enabled = False
        
    Case enumMonthlyReportType.AccountManagerCustPerf
        lblMonthlyRptCreator.Enabled = True
        cboMonthlyRptCreator.Enabled = True
        
    Case enumMonthlyReportType.AccountManagerCustAndProdPerf
        lblMonthlyRptCreator.Enabled = True
        cboMonthlyRptCreator.Enabled = True
        lblMonthlyRptOutletGrp.Enabled = True
        cboMonthlyRptOutletGrp.Enabled = True
        lblStartEndDate.Enabled = True
        cboStartEndDate.Enabled = True
        
End Select

Proc_Exit:
PopCallStack
Exit Sub

Err_Handler:
GlobalErrHandler
Resume Proc_Exit
End Sub

Private Sub SetMonthEndSelection()
Dim rs As ADODB.Recordset
Dim qry As String
Dim dte As String
Dim dteStart As Date
Dim dteEnd As Date
Dim dteTemp As Date

If gEnableErrorHandling Then On Error GoTo Err_Handler
PushCallStack "frmMonthlyReport|SetMonthEndSelection"

Set rs = New ADODB.Recordset

qry = "SELECT DISTINCT MonthDate FROM " & TRANS_TBL
rs.Open qry, cn, adOpenDynamic

' Set First month end date
rs.MoveFirst
dte = rs.Fields("MonthDate").Value
dteStart = CDate(WorksheetFunction.EoMonth("01-" & MonthName(Mid(dte, 5, 2)) & "-" & Mid(dte, 1, 4), 0))

' Set Last month end date
rs.MoveLast
dte = rs.Fields("MonthDate").Value
dteEnd = CDate(WorksheetFunction.EoMonth("01-" & MonthName(Mid(dte, 5, 2)) & "-" & Mid(dte, 1, 4), 0))

Call CloseRecordset(rs, True)

' Populate month end selection
dteTemp = dteStart
cboMonthEnd.Clear
Do Until dteTemp > dteEnd
    With cboMonthEnd
        .AddItem Format(dteTemp, "dd-mmm-yy")
    End With
    dteTemp = CDate(WorksheetFunction.EoMonth(DateAdd("d", 1, dteTemp), 0))
Loop

Proc_Exit:
PopCallStack
Exit Sub

Err_Handler:
GlobalErrHandler
Resume Proc_Exit
End Sub

Private Sub cmdViewMonthlyReport_Click()
If gEnableErrorHandling Then On Error GoTo Err_Handler
PushCallStack "frmMonthlyReport|cmdViewMonthlyReport_Click"

Call CreateMonthlyReport(cboMontlyReportType.Value, Me)

Proc_Exit:
PopCallStack
Exit Sub

Err_Handler:
GlobalErrHandler
Resume Proc_Exit
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If gEnableErrorHandling Then On Error GoTo Err_Handler
PushCallStack "frmMonthlyReport|UserForm_QueryClose"

If CloseMode <> 1 Then
    Cancel = 1
    Me.Hide
End If

Proc_Exit:
PopCallStack
Exit Sub

Err_Handler:
GlobalErrHandler
Resume Proc_Exit
End Sub
