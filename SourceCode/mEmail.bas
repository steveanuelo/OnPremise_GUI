Attribute VB_Name = "mEmail"
Option Explicit
Option Compare Text

Function SendEMail(Subject As String, _
        FromAddress As String, _
        ToAddress As String, _
        MailBody As String, _
        SMTP_Server As String, _
        BodyFileName As String, _
        Optional Attachments As Variant = Empty) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SendEmail Function
' By Chip Pearson, chip@cpearson.com www.cpearson.com 28-June-2012
'
' This function sends an email to the specified user.
' Parameters:
'   Subject:        The subject of the email.
'   FromAddress:    The sender's email address
'   ToAddress:      The recipient's email address or addresses.
'   MailBody:       The body of the email.
'   SMTP_Server:    The SMTP-Server name for outgoing mail.
'   BodyFileName:   A text file containing the body of the email.
'   Attachments     A single file name or an array of file names to
'                   attach to the message. The files must exist.
' Return Value:
'   True if successful.
'   False if failure.
'
' Subject may not be an empty string.
' FromAddress must be a valid email address.
' ToAddress must be a valid email address. To send to multiple recipients,
' use a semi-colon to separate the individual addresses. If there is a
' failure in one address, processing terminates and messages are not
' send to the rest of the recipients.
' If MailBody is vbNullString and BodyFileName is an existing text file, the content
' of the file named by BodyFileName is put into the body of the email. If
' BodyFileName does not exist, the function returns False. The content of
' the message body is created by a line-by-line import from BodyFileName.
' If MailBody is not vbNullString, then BodyFileName is ignored and the body
' is not created from the file.
' SMTP_Server must be a valid accessable SMTP server name.
' If both MailBody and BodyFileName are vbNullString, the mail message is
' sent with no body content.
' Attachments can be either a single file name as a String or an array of
' file names. If an attachment file does not exist, it is skipped but
' does not cause the procedure to terminate.
'
' If you want to send ThisWorkbook as an attachment to the message, use code
' like the following:
'    ThisWorkbook.Save
'    ThisWorkbook.ChangeFileAccess xlReadOnly
'    B = SendEmail( _
'        ... parameters ...
'        Attachments:=ThisWorkbook.FullName)
'    ThisWorkbook.ChangeFileAccess xlReadWrite
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Required References:
' --------------------
'   Microsoft CDO for Windows 2000 Library
'       Typical File Location: C:\Windows\system32\cdosys.dll
'       GUID: {CD000000-8B95-11D1-82DB-00C04FB1625D}
'       Major: 1    Minor: 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim MailMessage As CDO.Message
Dim N As Long
Dim FNum As Integer
Dim S As String
Dim Body As String
Dim Recips() As String
Dim Recip As String
Dim NRecip As Long

' ensure required parameters are present and valid.
If Len(Trim(Subject)) = 0 Then
    SendEMail = False
    Exit Function
End If

If Len(Trim(FromAddress)) = 0 Then
    SendEMail = False
    Exit Function
End If

If Len(Trim(SMTP_Server)) = 0 Then
    SendEMail = False
    Exit Function
End If

' Clean up the addresses
Recip = Replace(ToAddress, Space(1), vbNullString)
Recips = Split(Recip, ";")

For NRecip = LBound(Recips) To UBound(Recips)
    On Error Resume Next
    ' Create a CDO Message object.
    Set MailMessage = CreateObject("CDO.Message")
    If Err.Number <> 0 Then
        SendEMail = False
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
    With MailMessage
        .Subject = Subject
        .From = FromAddress
        .To = Recips(NRecip)
        If MailBody <> vbNullString Then
            .TextBody = MailBody
        Else
            If BodyFileName <> vbNullString Then
                If Dir(BodyFileName, vbNormal) <> vbNullString Then
                    ' import the text of the body from file BodyFileName
                    FNum = FreeFile
                    S = vbNullString
                    Body = vbNullString
                    Open BodyFileName For Input Access Read As #FNum
                    Do Until EOF(FNum)
                        Line Input #FNum, S
                        Body = Body & vbNewLine & S
                    Loop
                    Close #FNum
                    .TextBody = Body
                Else
                    ' BodyFileName not found.
                    SendEMail = False
                    Exit Function
                End If
            End If ' MailBody and BodyFileName are both vbNullString.
        End If
        
        If IsArray(Attachments) = True Then
            ' attach all the files in the array.
            For N = LBound(Attachments) To UBound(Attachments)
                ' ensure the attachment file exists and attach it.
                If Attachments(N) <> vbNullString Then
                    If Dir(Attachments(N), vbNormal) <> vbNullString Then
                        .AddAttachment Attachments(N)
                    End If
                End If
            Next N
        Else
            ' ensure the file exists and if so, attach it to the message.
            If Attachments <> vbNullString Then
                If Dir(CStr(Attachments), vbNormal) <> vbNullString Then
                    .AddAttachment Attachments
                End If
            End If
        End If
        With .Configuration.Fields
            ' set up the SMTP configuration
            .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP_Server
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
            .update
        End With
        
        On Error Resume Next
        Err.Clear
        ' Send the message
        .Send
        If Err.Number = 0 Then
            SendEMail = True
        Else
            SendEMail = False
            Exit Function
        End If
    End With
Next NRecip

End Function


Sub Test()

MsgBox SendEMail(Subject:="My Email", FromAddress:="steve_anuelo@lexiaconsulting.com", _
    ToAddress:="steve_anuelo@lexiaconsulting.com", MailBody:="Hello", _
    SMTP_Server:="smtp.gmail.com", BodyFileName:="Test for email sending via VBA")

End Sub
