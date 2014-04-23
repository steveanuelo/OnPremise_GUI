VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PromotionOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Promotions options class
'Use this class to create an object that contains all the possible selection options for a promo
'This class should only be needed in conjunction with a promo form

Option Explicit
Option Base 1

'Public prodOptions As ProductOptions
'Public custOptions As CustomerOptions
'Public freeGoodProdOptions As ProductOptions

Public arrDealTypes As Variant
Public arrRouteToMarket As Variant


Private Sub Class_Initialize()
'    Set prodOptions = New ProductOptions
'    Set custOptions = New CustomerOptions
'    Set freeGoodProdOptions = New ProductOptions
    
    ' Deal Types list
    ReDim arrDealTypes(1 To 3) As String
    arrDealTypes(1) = "Promotional Discount"
    arrDealTypes(2) = "Promotional Bonus Stock"
    arrDealTypes(3) = "Discount and Bonus stock"
    
    ' Route To Market list
    ReDim arrRouteToMarket(1 To 3) As String
    arrRouteToMarket(1) = "Indirect"
    arrRouteToMarket(2) = "Direct"
    arrRouteToMarket(3) = "Indirect and Direct"
    
End Sub

Public Sub update(promo As Promotion)
'    Me.custOptions.update promo.custGrp
'    Me.prodOptions.update promo.prdGrp
'    Me.freeGoodProdOptions.update promo.PromoDetails.freeGoodsPrdGrp
End Sub
