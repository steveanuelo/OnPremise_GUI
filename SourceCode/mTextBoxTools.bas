Attribute VB_Name = "mTextBoxTools"
Option Explicit

' Function to handle textbox edit toggle------------------Start
Public Sub ToggleTextBoxEdit(tbx1 As MSForms.TextBox, tbx2 As MSForms.TextBox, Optional tbx1a As MSForms.TextBox, Optional tbx2a As MSForms.TextBox)
    ' Both textbox are empty, enable both for editing
    If SetEmptyValue(tbx1, NullStrings) = vbNullString And SetEmptyValue(tbx2, NullStrings) = vbNullString Then
        Call EnableTextBox(tbx1)
        Call EnableTextBox(tbx2)
        
        If Not tbx1a Is Nothing Then Call EnableTextBox(tbx1a)
        If Not tbx2a Is Nothing Then Call EnableTextBox(tbx2a)
    End If
    
    ' If textbox 1 is not empty, then disable textbox 2
    If SetEmptyValue(tbx1, NullStrings) <> vbNullString And SetEmptyValue(tbx2, NullStrings) = vbNullString Then
        Call EnableTextBox(tbx1)
        Call EnableTextBox(tbx2, False)
        
        If Not tbx1a Is Nothing Then Call EnableTextBox(tbx1a)
        If Not tbx2a Is Nothing Then Call EnableTextBox(tbx2a, False)
    End If
    
    ' If textbox 2 is not empty, then disable textbox 1
    If SetEmptyValue(tbx2, NullStrings) <> vbNullString And SetEmptyValue(tbx1, NullStrings) = vbNullString Then
        Call EnableTextBox(tbx2)
        Call EnableTextBox(tbx1, False)
        
        If Not tbx2a Is Nothing Then Call EnableTextBox(tbx2a)
        If Not tbx1a Is Nothing Then Call EnableTextBox(tbx1a, False)
    End If
End Sub
' Function to handle textbox edit toggle------------------End

Public Sub EnableTextBox(tbx As MSForms.TextBox, Optional blnEnable As Boolean = True)
    tbx.Enabled = blnEnable
    If blnEnable Then
        tbx.BackColor = ENABLE_COLOR
    Else
        tbx.BackColor = DISABLE_COLOR
    End If
End Sub
