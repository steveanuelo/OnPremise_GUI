VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Promotion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Promotion class:
' This class holds all the information available for a given promotion
' A given promotion structure contains:
' - 1 set of Promotion Details(other promotion info)
' - 1 set of Deal Level info
' - 1 or more set of Product Discount info
' - 1 or more set of Product Bonus info


Option Explicit
Option Base 1

Public PromoDetails As PromotionDetails
'Public prdGrp As ProductGroup
'Public custGrp As CustomerGroup
'Public PromoResults As PromotionResults
Public ID As Long

Private Sub Class_Initialize()
'    Set prdGrp = New ProductGroup
'    Set custGrp = New CustomerGroup
    Set PromoDetails = New PromotionDetails
'    Set PromoResults = New PromotionResults
End Sub


