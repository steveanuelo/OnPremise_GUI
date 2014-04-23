VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_ExtendListBoxDimensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents ctrMainListBox As MSForms.ListBox
Attribute ctrMainListBox.VB_VarHelpID = -1
Private WithEvents ctrLabel As MSForms.Label
Attribute ctrLabel.VB_VarHelpID = -1
Private WithEvents ctrComboBox As MSForms.ComboBox
Attribute ctrComboBox.VB_VarHelpID = -1
Private WithEvents ctrTextBox As MSForms.TextBox
Attribute ctrTextBox.VB_VarHelpID = -1
Private WithEvents ctrListBox As MSForms.ListBox
Attribute ctrListBox.VB_VarHelpID = -1
Private WithEvents ctrCommandButton As MSForms.CommandButton
Attribute ctrCommandButton.VB_VarHelpID = -1
Private WithEvents ctrFrame As MSForms.Frame
Attribute ctrFrame.VB_VarHelpID = -1

Private dblOriginalHeight As Double
Private dblExtendedHeight As Double
Private dblOriginalWidth As Double
Private dblExtendedWidth As Double

Public Property Let OriginalHeight(dbl As Double)
    dblOriginalHeight = dbl
End Property

Public Property Let ExtendedHeight(dbl As Double)
    dblExtendedHeight = dbl
End Property

Public Property Let OriginalWidth(dbl As Double)
    dblOriginalWidth = dbl
End Property

Public Property Let ExtendedWidth(dbl As Double)
    dblExtendedWidth = dbl
End Property

Public Property Set MainListBox(objNew As MSForms.ListBox)
    Set ctrMainListBox = objNew
End Property

Public Property Set Label(objNew As MSForms.Label)
    Set ctrLabel = objNew
End Property

Public Property Set ComboBox(objNew As MSForms.ComboBox)
    Set ctrComboBox = objNew
End Property

Public Property Set TextBox(objNew As MSForms.TextBox)
    Set ctrTextBox = objNew
End Property

Public Property Set ListBox(objNew As MSForms.ListBox)
    Set ctrListBox = objNew
End Property

Public Property Set CommandButton(objNew As MSForms.CommandButton)
    Set ctrCommandButton = objNew
End Property

Public Property Set Frame(objNew As MSForms.Frame)
    Set ctrFrame = objNew
End Property

Private Sub ctrMainListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    With ctrMainListBox
        If .Height <> dblExtendedHeight Then .Height = dblExtendedHeight
        If .Width <> dblExtendedWidth Then .Width = dblExtendedWidth
        '.SetFocus
        .ZOrder 0
    End With
End Sub

Private Sub ctrLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call ResetlstDimensions
End Sub

Private Sub ctrComboBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call ResetlstDimensions
End Sub

Private Sub ctrTextBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call ResetlstDimensions
End Sub

Private Sub ctrListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call ResetlstDimensions
End Sub

Private Sub ctrCommandButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call ResetlstDimensions
End Sub

Private Sub ctrFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call ResetlstDimensions
End Sub

Private Sub ResetlstDimensions()
    With ctrMainListBox
        .Height = dblOriginalHeight
        .Width = dblOriginalWidth
    End With
End Sub

Private Sub Class_Terminate()
    Set ctrMainListBox = Nothing
    Set ctrLabel = Nothing
    Set ctrComboBox = Nothing
    Set ctrTextBox = Nothing
    Set ctrListBox = Nothing
    Set ctrCommandButton = Nothing
    Set ctrFrame = Nothing
End Sub
