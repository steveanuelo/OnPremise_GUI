Attribute VB_Name = "mCreateMinMax"
Option Explicit

Private Declare Function FindWindowA Lib "user32" _
(ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long
 
Private Declare Function GetWindowLongA Lib "user32" _
(ByVal hwnd As Long, _
ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLongA Lib "user32" _
(ByVal hwnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Public Sub FormatUserForm(UserFormCaption As String)
Dim hwnd As Long
Dim exLong As Long
    
If gEnableErrorHandling Then On Error GoTo Err_Handler
PushCallStack "mCreateMinMax|FormatUserForm"

hwnd = FindWindowA(vbNullString, UserFormCaption)
exLong = GetWindowLongA(hwnd, -16)
If (exLong And &H20000) = 0 Then
    SetWindowLongA hwnd, -16, exLong Or &H20000
Else
End If
    
Proc_Exit:
PopCallStack
Exit Sub

Err_Handler:
GlobalErrHandler
Resume Proc_Exit
End Sub
