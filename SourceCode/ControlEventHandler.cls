VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
'
'Public WithEvents TextBox As MSForms.TextBox
'Public WithEvents ListBox As MSForms.ListBox
'Public WithEvents ComboBox As MSForms.ComboBox
'
'Private Sub ListBox_Change()
'    If Not g_blnUpdateInProgress And Not IsNull(Me.ListBox.Value) Then Call processChangeEvents(Me.ListBox.Name, Me.ListBox.Value)
'End Sub
'
'Private Sub TextBox_Change()
'    'Textboxes don't have dependents so no need to update options
'    If Not g_blnUpdateInProgress Then
'        g_blnUpdateInProgress = True
'        Dim lngError As Long
'        lngError = frmPromo.updatePromo(Me.TextBox.Name, Me.TextBox.Value)
'        lngError = frmPromo.updateControls
'        g_blnUpdateInProgress = False
'    End If
'End Sub
'
'
'Private Sub ComboBox_Change()
'    If Not g_blnUpdateInProgress Then Call processChangeEvents(Me.ComboBox.Name, Me.ComboBox.Value)
'End Sub
'
'Public Sub processChangeEvents(strName As String, strValue As String)
'
''    g_blnUpdateInProgress = True
''
''    Dim lngError As Long
''
''    lngError = frmPromo.updatePromo(strName, strValue)
''    frmPromo.promoOptions.update frmPromo.promo
''    lngError = frmPromo.updateControls
''
''    g_blnUpdateInProgress = False
'
'End Sub
'
'
''Private Sub TextBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
''    If Not g_blnUpdateInProgress Then
''        If Right(Me.TextBox.Name, 4) = "Date" Then
''            g_bForm = True
''            frmCalendar.Show_Cal
''            Me.TextBox.Value = g_sDate
''        End If
''    End If
''End Sub
