VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_StandardDealSheetRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public rngRefNum As String
Public rngPRAContactName As String
Public rngPRAContactPhone As String
Public rngPRAContactEmail As String
Public rngPRAAuthoriserName As String
Public rngPRAAuthoriserEmail As String
Public rngCustomerContactName As String
Public rngCustomerContactPhone As String
Public rngCustomerContactEmail As String
Public rngBanner As String
Public rngBannerRegionName As String
Public rngOutletName As String
Public rngOutetNumber As String
Public rngState_All As String
Public rngState_NSW_ACT As String
Public rngState_VIC_TAS As String
Public rngState_QLD As String
Public rngState_WA As String
Public rngState_SA As String
Public rngState_NT As String
Public rngPromoStartDate As String
Public rngPromoEndDate As String
Public rngBuyingPeriodStartDate As String
Public rngBuyingPeriodEndDate As String
Public rngPerCustInGrpMixBuy As String
Public rngPerBannRegMixBuy As String
Public rngPerCustInGrpNumCases As String
Public rngPerBannRegNumCases As String
Public rngComments As String
Public rngProductInfo As String


Private Sub Class_Initialize()
    rngRefNum = "C3"
    rngPRAContactName = "C7"
    rngPRAContactPhone = "C8"
    rngPRAContactEmail = "C9"
    rngPRAAuthoriserName = "C10"
    rngPRAAuthoriserEmail = "C11"
    rngCustomerContactName = "C14"
    rngCustomerContactPhone = "C15"
    rngCustomerContactEmail = "C16"
    rngBanner = "C21"
    rngBannerRegionName = "C22"
    rngOutletName = "C23"
    rngOutetNumber = "C24"
    rngState_All = "I21"
    rngState_NSW_ACT = "I22"
    rngState_VIC_TAS = "I23"
    rngState_QLD = "I24"
    rngState_WA = "K22"
    rngState_SA = "K23"
    rngState_NT = "K24"
    rngPromoStartDate = "C27"
    rngPromoEndDate = "E27"
    rngBuyingPeriodStartDate = "C28"
    rngBuyingPeriodEndDate = "E28"
    rngPerCustInGrpMixBuy = "I27"
    rngPerBannRegMixBuy = "I28"
    rngPerCustInGrpNumCases = "J27"
    rngPerBannRegNumCases = "J28"
    rngComments = "M4"
    rngProductInfo = "B39"
End Sub

