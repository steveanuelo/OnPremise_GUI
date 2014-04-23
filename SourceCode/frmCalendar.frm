VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendar 
   Caption         =   "Calendar"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3810
   OleObjectBlob   =   "frmCalendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Buttons()  As New CalendarComButton

Sub Show_Cal(Optional strDate As String = "")
    'use class module to create commandbutton collection, then show calendar

    Dim iCmdBtns As Integer
    Dim ctl    As Control

    iCmdBtns = 0
    For Each ctl In frmCalendar.Controls
        If TypeName(ctl) = "CommandButton" Then
            If ctl.Name <> "CB_Close" Then
                iCmdBtns = iCmdBtns + 1
                ReDim Preserve Buttons(1 To iCmdBtns)
                Set Buttons(iCmdBtns).CmdBtnGroup = ctl
            End If
        End If
    Next ctl
    
    If Len(strDate) <> 0 Then
        frmCalendar.CB_Mth.Text = MonthName(Month(CDate(strDate)))
        frmCalendar.CB_Yr.Text = Year(CDate(strDate))
    End If
        
    frmCalendar.Show
End Sub

Private Sub CB_Close_Click()
    Unload Me
End Sub

Private Sub D19_Click()
    addDate
End Sub
Sub addDate()
    ActiveCell.Value = Parent
End Sub

Private Sub UserForm_Initialize()

    Dim i      As Long
    Dim lYearsAdd As Long
    Dim lYearStart As Long

    lYearStart = Year(Date) - 10
    lYearsAdd = Year(Date) + 10
    With Me
        For i = 1 To 12
            .CB_Mth.AddItem Format(DateSerial(Year(Date), i, 1), "mmmm")
        Next

        For i = lYearStart To lYearsAdd
            .CB_Yr.AddItem Format(DateSerial(i, 1, 1), "yyyy")
        Next

        .Tag = "Calendar"
        .CB_Mth.ListIndex = Month(Date) - 1
        .CB_Yr.ListIndex = Year(Date) - lYearStart
        .Tag = ""
    End With
    Call Build_Calendar

End Sub

Private Sub CB_Mth_Change()
    If Not Me.Tag = "Calendar" Then Build_Calendar
End Sub

Private Sub CB_Yr_Change()
    If Not Me.Tag = "Calendar" Then Build_Calendar
End Sub

Sub Build_Calendar()

    Dim i      As Integer
    Dim dTemp  As Date
    Dim dTemp2 As Date
    Dim iFirstDay As Integer
    With Me
        .Caption = " " & .CB_Mth.Value & " " & .CB_Yr.Value

        dTemp = CDate("01/" & .CB_Mth.Value & "/" & .CB_Yr.Value)
        iFirstDay = WeekDay(dTemp, vbSunday)
        .Controls("D" & iFirstDay).SetFocus

        For i = 1 To 42
            With .Controls("D" & i)
                dTemp2 = DateAdd("d", (i - iFirstDay), dTemp)
                .Caption = Format(dTemp2, "d")
                .Tag = dTemp2
                .ControlTipText = Format(dTemp2, "dd/mm/yy")
                'add dates to the buttons
                If Format(dTemp2, "mmmm") = CB_Mth.Value Then
                    If .BackColor <> &H80000016 Then .BackColor = &H80000018
                    If Format(dTemp2, "m/d/yy") = Format(Date, "m/d/yy") Then .SetFocus
                    .Font.Bold = True
                Else
                    If .BackColor <> &H80000016 Then .BackColor = &H8000000F
                    .Font.Bold = False
                End If
                'format the buttons
            End With
        Next
    End With

End Sub
