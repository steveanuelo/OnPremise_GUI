Attribute VB_Name = "mDataBase"
Option Explicit

Public cn As ADODB.Connection
Public cnRemote As ADODB.Connection

Public Const DB_NAME = "OP_Mappings.accdb"
Public Const REMOTE_DB_LOCATION = "I:\01_Shared_Data\Commercial Support\Lexia Gui Database\OP gui\"
'Public Const REMOTE_DB_LOCATION = "C:\Users\Public\Documents\Projects\AUS COOP\Remote DB\"

Public Function SetDBConnection(conn As ADODB.Connection, Optional blnIsLocalDB As Boolean = True) As Integer
      Dim Cnct As String
      Dim DBFullName As String

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mDataBase|SetDBConnection"

3     If blnIsLocalDB Then
4         DBFullName = ThisWorkbook.Path & "\" & DB_NAME
5     Else
6         DBFullName = REMOTE_DB_LOCATION & DB_NAME
7     End If

      ' Set connection
8     If conn Is Nothing Then
9         Set conn = New ADODB.Connection
10        Cnct = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                 "Data Source=" & DBFullName & ";" & _
                 "Jet OLEDB:Database " & Chr(80) & Chr(97) & Chr(115) & _
                 Chr(115) & Chr(119) & Chr(111) & Chr(114) & Chr(100) & _
                 Chr(61) & Chr(112) & Chr(114) & Chr(97) & Chr(117) & Chr(36) & ";"
                 
11        conn.Open ConnectionString:=Cnct
12    End If

Proc_Exit:
13    PopCallStack
14    Exit Function

Err_Handler:
15    GlobalErrHandler
16    Resume Proc_Exit
End Function

Public Sub CloseDBConnection(conn As ADODB.Connection)
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mDatabase|CloseDBConnection"

3     If Not (conn Is Nothing) Then
4         If conn.State <> adStateClosed Then
5             conn.Close
6         End If
7     End If
8     Set conn = Nothing

Proc_Exit:
9     PopCallStack
10    Exit Sub

Err_Handler:
11    GlobalErrHandler
12    Resume Proc_Exit
End Sub

Public Sub CloseRecordset(rs As ADODB.Recordset, Optional blnSetToNothing As Boolean = False)
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mDataBase|CloseRecordset"

3     If Not (rs Is Nothing) Then
4         If rs.State <> adStateClosed Then
5             rs.Close
6         End If
7     End If

8     If blnSetToNothing = True Then
9         Set rs = Nothing
10    End If

Proc_Exit:
11    PopCallStack
12    Exit Sub

Err_Handler:
13    GlobalErrHandler
14    Resume Proc_Exit
End Sub

Public Function IsItemExistInTable(tbl As String, fld As String, strItem As String, _
                                   Optional strValidation As String = vbNullString, _
                                   Optional blnConnLocal As Boolean = True) As Boolean
          Dim conn As ADODB.Connection
          Dim rs As ADODB.Recordset
          
1         IsItemExistInTable = False
          
2         If blnConnLocal = True Then
3             Set conn = cn
4         Else
5             Set conn = cnRemote
6         End If
          
7         Set rs = New ADODB.Recordset
8         With rs
9             .Open "SELECT DISTINCT " & fld & " FROM " & tbl & _
                      " WHERE " & fld & " = " & strValidation & strItem & strValidation & ";", conn
10            If Not .EOF Then
11                IsItemExistInTable = True
12            End If
13            .Close
14        End With
          
15        Set rs = Nothing
End Function

Public Function GetItemFromMappingTbl(tbl As String, strReturnField As String, _
                                      Optional strWhereField As String, Optional strSearchValue As String, _
                                      Optional strValidation As String = vbNullString, _
                                      Optional blnConnLocal As Boolean = True, _
                                      Optional strWhereCondit As String = "") As String
          Dim conn As ADODB.Connection
          Dim rs As ADODB.Recordset
          Dim qry As String
          
1         If blnConnLocal = True Then
2             Set conn = cn
3         Else
4             Set conn = cnRemote
5         End If
          

6         qry = "SELECT DISTINCT " & strReturnField & " " & _
                "FROM " & tbl & " "
7         If Len(strWhereCondit) = 0 Then
8             qry = qry & "WHERE " & strWhereField & " = " & strValidation & strSearchValue & strValidation & ";"
9         Else
10            qry = qry & "WHERE " & strWhereCondit
11        End If
          
          
12        Set rs = New ADODB.Recordset
13        With rs
14            .Open qry, conn
15            If Not .EOF Then
16                If Len(rs.Fields(0).Value) <> 0 Then
17                    GetItemFromMappingTbl = CStr(rs.Fields(0).Value)
18                Else
19                    GetItemFromMappingTbl = vbNullString
20                End If
21            End If
22            .Close
23        End With
          
24        Set rs = Nothing
End Function

Public Function GetItemFromMappingTbl2(tbl As String, strReturnField As String, _
                                       strWhereCondition As String, _
                                       Optional blnConnLocal As Boolean = True) As String
          Dim conn As ADODB.Connection
          Dim rs As ADODB.Recordset
          Dim qry As String
          
1         If blnConnLocal = True Then
2             Set conn = cn
3         Else
4             Set conn = cnRemote
5         End If
          
6         qry = "SELECT DISTINCT " & strReturnField & " " & _
                "FROM " & tbl & " " & _
                "WHERE " & strWhereCondition
          
7         Set rs = New ADODB.Recordset
8         With rs
9             .Open qry, conn
10            If Not .EOF Then
11                If Len(rs.Fields(0).Value) <> 0 Then
12                    GetItemFromMappingTbl = CStr(rs.Fields(0).Value)
13                Else
14                    GetItemFromMappingTbl = vbNullString
15                End If
16            End If
17            .Close
18        End With
          
19        Set rs = Nothing
End Function

'
'
'Public Function GetArrayList(qry As String, Optional blnMultiColumn As Boolean = False) As Variant
'
'    Dim arr1 As Variant
'    Dim arr2 As Variant
'    Dim x As Long, y As Long
'
'
'    ' Database information
'    DBFullName = ThisWorkbook.Path & "\Mappings.accdb"
'
'    ' Open the connection
'    Set cn = New ADODB.Connection
'    Cnct = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
'           "Data Source=" & DBFullName & ";"
'    cn.Open ConnectionString:=Cnct
'
'    ' Create RecordSet
'    Set rs = New ADODB.Recordset
'    With rs
''        ' Filter
''        qry = "SELECT DISTINCT Banner FROM " & CUSTOMER_MAP_TBL & ";"
'
'        .Open qry, ActiveConnection:=cn
'        If Not .EOF Then
'            ' Write the recordset
'            'Worksheets("Temp").Range("A1").Offset(1, 0).CopyFromRecordset rs
'            arr1 = .GetRows(100000)
'
'            Select Case blnMultiColumn
'                Case False
'                    ReDim arr2(UBound(arr1, 2))
'                    For x = 0 To UBound(arr1, 2)
'                        arr2(x) = arr1(0, x)
'                    Next x
'                Case True
'                    ReDim arr2(UBound(arr1, 2), UBound(arr1))
'                    For x = 0 To UBound(arr1)
'                        For y = 0 To UBound(arr1, 2)
'                            arr2(y, x) = arr1(x, y)
'                        Next y
'                    Next x
'            End Select
'
'
'
'            GetArrayList = arr2
'        End If
'        .Close
'    End With
'
'    Set rs = Nothing
'    cn.Close
'    Set cn = Nothing
'End Function
'
