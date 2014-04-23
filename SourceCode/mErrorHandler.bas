Attribute VB_Name = "mErrorHandler"
Option Explicit

' Current pointer to the array element of the call stack
Private mintStackPointer As Integer

' Array of procedure names in the call stack
Private mastrCallStack() As String

' The number of elements to increase the array
Private Const mcintIncrementStackSize As Integer = 10

' Add the current procedure name to the Call Stack.
' Should be called whenever a procedure is called
Public Sub PushCallStack(strProcName As String)
    On Error Resume Next
    
    ' Verify the stack array can handle the current array element
    If mintStackPointer > UBound(mastrCallStack) Then
        ' If array has not been defined, initialize the error handler
        If Err.Number = 9 Then
            ErrorHandlerInit
        Else
          ' Increase the size of the array to not go out of bounds
            ReDim Preserve mastrCallStack(UBound(mastrCallStack) + _
              mcintIncrementStackSize)
        End If
    End If
    
    On Error GoTo 0
    
    mastrCallStack(mintStackPointer) = strProcName
    
    ' Increment pointer to next element
    mintStackPointer = mintStackPointer + 1
End Sub

Private Sub ErrorHandlerInit()
    mintStackPointer = 1
    ReDim mastrCallStack(1 To mcintIncrementStackSize)
End Sub
' Remove a procedure name from the call stack
Sub PopCallStack()
    If mintStackPointer <= UBound(mastrCallStack) Then
        mastrCallStack(mintStackPointer) = ""
    End If
    
    ' Reset pointer to previous element
    mintStackPointer = mintStackPointer - 1
End Sub

' Main procedure to handle errors that occur.
Public Sub GlobalErrHandler()
    Dim strError As String
    Dim lngError As Long
    Dim intErl As Integer
    Dim strMsg As String
    
    ' Variables to preserve error information
    strError = Err.Description
    lngError = Err.Number
    intErl = Erl
    
    ' Prompt the user with information on the error:
    strMsg = "Module: " & Split(CurrentProcName, "|")(0) & vbCrLf & _
             "Procedure: " & Split(CurrentProcName, "|")(1) & vbCrLf & _
             "Line : " & intErl & vbCrLf & _
             "Error : (" & lngError & ") " & strError
    MsgBox strMsg, vbCritical
    
    ' Write error to file:
    WriteErrorToFile lngError, strError, intErl
    
    ' Reset workspace, close open objects
    Call ResetWorkspace
    
    ' Exit Access without saving any changes
    ' (you may want to change this to save all changes)
    End
End Sub

Private Function CurrentProcName() As String
    CurrentProcName = mastrCallStack(mintStackPointer - 1)
End Function

Private Sub ResetWorkspace()
    Dim intCounter As Integer

    On Error Resume Next

    ' Close database connections
    Call CloseDBConnection(cn)

    ' Enable alert messages
    Application.DisplayAlerts = True
    ' Set mouse pointer to default arrow
    DoEvents

    ' Unload all open forms?
    DoEvents

End Sub

Public Sub WriteErrorToFile(lngError As Long, strError As String, intErl As Integer)
    Dim intFileNum As Integer
    Dim i As Integer
    Const ERROR_LOG_FILE = "ErrorLog.txt"

    
    intFileNum = FreeFile
    Open ThisWorkbook.Path & "\" & ERROR_LOG_FILE For Append As #intFileNum    ' Open file for output.
    
    Print #intFileNum, Format(Now, "yyyy.mm.dd hh:mm"); vbCrLf; _
        "Error Number: "; lngError; vbCrLf; _
        "Error Description: "; strError; vbCrLf; _
        "Call Stack: "; vbCrLf; _
        GetCallStack; _
        "Error Line in Code: "; intErl; vbCrLf; vbCrLf
    
    Close #intFileNum
End Sub

Private Function GetCallStack() As String
    Dim i As Integer
    Dim str As String
    
    str = vbNullString
    For i = 1 To UBound(mastrCallStack)
        If Len(mastrCallStack(i)) <> 0 Then
            str = str & mastrCallStack(i) & vbCrLf
        End If
    Next i
    
    GetCallStack = str
End Function

' Sample procedure with error handler
'Sub AdvancedErrorStructure()
'      ' Use a call stack and global error handler
'
'      If gcfHandleErrors Then On Error GoTo PROC_ERR
'      PushCallStack "AdvancedErrorStructure"
'
'      ' << Your code here >>
'
'PROC_EXIT:
'    PopCallStack
'    Exit Sub
'
'PROC_ERR:
'    GlobalErrHandler
'    Resume PROC_EXIT
'End Sub

Sub SampleErrorWithLineNumbers()
Dim dblNum As Double

If gEnableErrorHandling Then On Error GoTo Err_Handler
PushCallStack "mErrorHandler|SampleErrorWithLineNumbers"

Select Case Rnd()
  Case Is < 0.2
    dblNum = 5 / 0
  Case Is < 0.4
    dblNum = 5 / 0
  Case Is < 0.6
    dblNum = 5 / 0
  Case Is < 0.8
    dblNum = 5 / 0
  Case Else
End Select
    
Proc_Exit:
PopCallStack
Exit Sub

Err_Handler:
GlobalErrHandler
Resume Proc_Exit

End Sub

