Attribute VB_Name = "mExcelFormatting"
Option Explicit

Public Sub FormatBorder(rng As Range, _
                        Optional FillColor As Long = xlColorIndexNone, _
                        Optional TopStyle As XlLineStyle = xlNone, _
                        Optional TopWeight As XlBorderWeight = xlHairline, _
                        Optional BottomStyle As XlLineStyle = xlNone, Optional BottomWeight As XlBorderWeight = xlHairline, _
                        Optional LeftStyle As XlLineStyle = xlNone, Optional LeftWeight As XlBorderWeight = xlHairline, _
                        Optional RightStyle As XlLineStyle = xlNone, Optional RightWeight As XlBorderWeight = xlHairline, _
                        Optional InsideHorizontalStyle As XlLineStyle = xlNone, Optional InsideHorizontalWeight As XlBorderWeight = xlHairline, _
                        Optional InsideVerticalStyle As XlLineStyle = xlNone, Optional InsideVerticalWeight As XlBorderWeight = xlHairline)
    
    ' Top Border
    With rng.Borders(xlEdgeTop)
        If TopStyle <> xlNone Then
            .LineStyle = TopStyle
            .Weight = TopWeight
            .ColorIndex = 0
        End If
    End With

    ' Bottom Border
    With rng.Borders(xlEdgeBottom)
        If BottomStyle <> xlNone Then
            .LineStyle = BottomStyle
            .Weight = BottomWeight
            .ColorIndex = 0
        End If
    End With
    
    ' Left Border
    With rng.Borders(xlEdgeLeft)
        If LeftStyle <> xlNone Then
            .LineStyle = LeftStyle
            .Weight = LeftWeight
            .ColorIndex = 0
        End If
    End With
    
    ' Right Border
    With rng.Borders(xlEdgeRight)
        If RightStyle <> xlNone Then
            .LineStyle = RightStyle
            .Weight = RightWeight
            .ColorIndex = 0
        End If
    End With

    ' Inside Horizontal Border
    With rng.Borders(xlInsideHorizontal)
        If InsideHorizontalStyle <> xlNone Then
            .LineStyle = InsideHorizontalStyle
            .Weight = InsideHorizontalWeight
            .ColorIndex = 0
        End If
    End With

    ' Inside Vertical Border
    With rng.Borders(xlInsideVertical)
        If InsideVerticalStyle <> xlNone Then
            .LineStyle = InsideVerticalStyle
            .Weight = InsideVerticalWeight
            .ColorIndex = 0
        End If
    End With
    
    ' Fill color
    With rng.Interior
        If FillColor <> xlColorIndexNone Then
            .Color = FillColor
            .TintAndShade = 0
            .Pattern = xlSolid
        End If
    End With

End Sub



