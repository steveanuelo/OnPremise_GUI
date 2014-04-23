VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PromotionDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' This class contains the structure of the promotion info details

Public promoRefNumber As String
Public promoCreator As String
Public promoManager As String

Public customerName As String
Public customerPhoneNumber As Integer
Public customerEmail As String

Public dealType As String

Public startDate As Date
Public endDate As Date
Public buyingPeriodStartDate As Date
Public buyingPeriodEndDate As Date

Public routeToMarket As String
Public wholesaler As Variant        'array?

Public commentsPromoParameters As String

Public isMixedBuyPerCustomerInGroup As Boolean
Public isMixedBuyPerBannerRegion As Boolean
Public numberOfCasesPerCustomerInGroup As Integer
Public numberOfCasesPerBannerRegion As Integer

Public commentsProductDetails As String


Public Property Get buyingStartDate() As Date
    'buyingStartDate = DateDiff("ww", startDate - 2)
End Property

'Public Property Let buyingStartDate(dtm As Date)
'    buyingPeriodStartDate = dtm
'End Property

Public Property Get buyingEndDate() As Date
'    buyingEndDate = DateAdd("ww", endDate + 1)
End Property

'Public Property Let buyingPeriodEndDate(dtm As Date)
'    buyingPeriodEndDate = dtm
'End Property





