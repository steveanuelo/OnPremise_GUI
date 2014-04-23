VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_ALMDealSheetRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public rngRefNum As String
Public rngPromoStartDate As String
Public rngPromoEndDate As String
Public rngBuyingPeriodStartDate As String
Public rngBuyingPeriodEndDate As String
Public rngCustName As String        ' Outet Name
Public rngCustNumber As String      ' Outlet Number
Public rngSubmittedBy As String     ' PRA Promo Creator
Public rngState_NSW As String
Public rngState_VIC As String
Public rngState_QLD As String
Public rngState_SA As String
Public rngState_WA As String
Public rngState_TAS As String
Public rngState_NT As String
Public rngState_ACT As String
Public rngGroupName As String       ' Banner or Banner Region level
Public rngQtyRestrictionNumCases As String
Public rngProductInfo As String
Public rngInsertPoint As String

Private Sub Class_Initialize()
    rngRefNum = "F2"
    rngPromoStartDate = "F5"
    rngPromoEndDate = "G5"
    rngBuyingPeriodStartDate = "D8"
    rngBuyingPeriodEndDate = "D9"
    rngCustName = "G8"
    rngCustNumber = "G9"
    rngSubmittedBy = "P10"
    rngState_NSW = "C12"
    rngState_VIC = "C13"
    rngState_QLD = "C14"
    rngState_SA = "C15"
    rngState_WA = "E12"
    rngState_TAS = "E13"
    rngState_NT = "E14"
    rngState_ACT = "E15"
    rngGroupName = "G13"
    rngQtyRestrictionNumCases = "S15"
    rngProductInfo = "E25"
    
    rngInsertPoint = "A45"
End Sub
