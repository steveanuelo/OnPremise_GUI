VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalendarComButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'use a class module to create a collection of commandbuttons

Public WithEvents CmdBtnGroup As MSForms.CommandButton
Attribute CmdBtnGroup.VB_VarHelpID = -1

Sub CmdBtnGroup_Click()
    If Month(CDate(CmdBtnGroup.Tag)) <> frmCalendar.CB_Mth.ListIndex + 1 Then
        Select Case _
               MsgBox("The selected date is not in the currently selected month." _
                      & vbNewLine & "Continue?", _
                      vbYesNo Or vbExclamation Or vbDefaultButton1, "Date check")
            Case vbYes
                If g_bForm Then
                    GoTo on_Form
                Else: GoTo addDate
                End If
            Case vbNo
                Exit Sub
        End Select
    Else:
    If g_bForm Then
        GoTo on_Form
    Else: GoTo addDate
    End If
addDate:
    With ActiveCell
        .Value = CDate(CmdBtnGroup.Tag)
        .EntireColumn.AutoFit
    End With
    GoTo chg_month
on_Form:
    g_sDate = CmdBtnGroup.Tag
chg_month:
    With frmCalendar.CB_Mth
        .ListIndex = Month(CmdBtnGroup.Tag) - 1
    End With
     End If
    Unload frmCalendar
End Sub


