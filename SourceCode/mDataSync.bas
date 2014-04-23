Attribute VB_Name = "mDataSync"
Option Explicit

Public Const SYNC_DATE_TBL = "T_Last_Sync_Date"

Public Enum enumSyncType
    promo = 0
    COOP = 1
End Enum

Public Sub SyncDatabase()
      Dim strUserID As String
      Dim strUserPermission As String
      Dim dteLastSyncLocal As Date
      Dim dteLastSyncRemote As Date
      Dim qry As String
      Dim blnToUpdate As Boolean
      Dim arrMainData As Variant
      Dim arrMainTbl As Variant
      Dim arrOtherLocal As Variant
      Dim arrOtherTbl As Variant
      Dim arrOpOtherTbl As Variant
      Dim arrStaticTbl As Variant
      Dim i As Integer, j As Integer, x As Integer, y As Integer, z As Integer
      Dim rs As ADODB.Recordset
      Dim strPWD As String

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mDataSync|SyncDatabase"

      ' Set Main tables
3     arrMainTbl = Array(OP_MAIN_TBL)

      ' Set other tables
4     arrOpOtherTbl = Array(OP_PROD_DETAILS_TBL, OP_PROD_NON_QA3_TBL)

5     Set rs = New ADODB.Recordset

      ' Get UserID
6     strUserID = GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "ID", "WinLoginName", UCase(Environ("UserName")), """")

      ' Get user permission (admin or ordinary user)
7     strUserPermission = GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "AccessTypeID", "WinLoginName", UCase(Environ("UserName")), """")

      ' Get date of last sync of user
8     dteLastSyncLocal = CDate(GetItemFromMappingTbl(SYNC_DATE_TBL, "LastSyncDate", "ID", strUserID, """"))
         
      ' Open a connection to the remote db
9     Call SetDBConnection(cnRemote, False)

      ' Get date of last sync of user in the remote db
10    dteLastSyncRemote = CDate(GetItemFromMappingTbl(SYNC_DATE_TBL, "LastSyncDate", "ID", strUserID, """", False))

      ' Have new updated data from local db
11    If dteLastSyncLocal > dteLastSyncRemote Then
          
          ' Loop through COOP tables
12        For i = 0 To UBound(arrMainTbl)
                 
              ' Get only records that are later than the last sync remote date
13            qry = vbNullString
14            qry = qry & "SELECT T1.* " & _
                          "FROM " & arrMainTbl(i) & " AS T1 INNER JOIN " & PRA_EMPLOYEE_TBL & " AS T2 ON T1.CreatorID = T2.ID " & _
                          "WHERE T1.[LastSyncDate] > #" & dteLastSyncRemote & "# "
              ' Get every record that are owned by the user
              ' else get every later records if user has admin permission
15            If strUserPermission = enumUserPermission.OrdinaryUser Then
16                qry = qry & "AND T2.ID = '" & strUserID & "'"
17            End If
              
18            arrMainData = GetArrayList(qry, True)
              
19            If IsArrayAllocated(arrMainData) Then
                  ' Loop through all the local data
20                For j = 0 To UBound(arrMainData)
                      ' Flag to update the other tables
21                    blnToUpdate = False
                      
                      ' Check if Ref# is already existing.
                      ' If yes update, else add a new record
22                    If IsItemExistInTable(CStr(arrMainTbl(i)), "RefNumber", CStr(arrMainData(j, 0)), "'", False) Then
                          ' Update
                          ' Check if the local sync date is later than the remote sync date
23                        If CDate(arrMainData(j, UBound(arrMainData, 2))) > _
                             CDate(GetItemFromMappingTbl(CStr(arrMainTbl(i)), "LastSyncDate", "RefNumber", CStr(arrMainData(j, 0)), "'", False)) Then
                              
                              ' Open recordset to enable changes
24                            qry = "SELECT * FROM " & arrMainTbl(i) & " " & _
                                    "WHERE RefNumber='" & arrMainData(j, 0) & "'"
25                            rs.Open qry, cnRemote, adOpenKeyset, adLockOptimistic, adCmdText
                              
                              ' Change data in edit buffer
                              ' Do not update Ref#, Creator, Authoriser
26                            For x = 3 To UBound(arrMainData, 2)
27                                rs.Fields(x).Value = arrMainData(j, x)
28                            Next x
29                            rs.update
30                            Call CloseRecordset(rs)
                              
                              ' When successfully updated
31                            blnToUpdate = True
32                        Else    ' No updates done
                              ' Store table, ref# to the not updated Array
33                            DoEvents
34                        End If
35                    Else
                          ' Append
                          ' Open table to enable appending
36                        rs.Open arrMainTbl(i), cnRemote, adOpenKeyset, adLockOptimistic, adCmdTable
                          
                          ' Add new record
37                        rs.AddNew
38                        For x = 0 To UBound(arrMainData, 2)
39                            rs.Fields(x).Value = arrMainData(j, x)
40                        Next x
41                        rs.update
42                        Call CloseRecordset(rs)
                          
43                        blnToUpdate = True
44                    End If
                      
                      ' Update other tables if the update or append are successfull
45                    If blnToUpdate Then
                          '
46                        arrOtherTbl = arrOpOtherTbl
                          
47                        For x = 0 To UBound(arrOtherTbl)
                              ' Get other local table records with the Ref#
48                            qry = vbNullString
49                            qry = qry & "SELECT T1.* " & _
                                          "FROM " & arrOtherTbl(x) & " AS T1 " & _
                                          "WHERE T1.RefNumber = '" & CStr(arrMainData(j, 0)) & "' "
50                            arrOtherLocal = GetArrayList(qry, True)
                              
51                            If IsArrayAllocated(arrOtherLocal) Then
                              
                                  ' Delete from remote
52                                qry = vbNullString
53                                qry = "DELETE * " & _
                                        "FROM " & arrOtherTbl(x) & " " & _
                                        "WHERE RefNumber = '" & CStr(arrMainData(j, 0)) & "';"
54                                cnRemote.Execute qry
                          
                                  ' Append to remote
                                  ' Open table to enable appending
55                                rs.Open arrOtherTbl(x), cnRemote, adOpenKeyset, adLockOptimistic, adCmdTable
                                  
                                  ' Add new record
56                                For z = 0 To UBound(arrOtherLocal)
57                                    rs.AddNew
                                      
58                                    For y = 0 To UBound(arrOtherLocal, 2)
59                                        rs.Fields(y).Value = arrOtherLocal(z, y)
60                                    Next y
                                      
61                                    rs.update
62                                Next z
                                  
63                                Call CloseRecordset(rs)
                              
64                            End If
                              
65                        Next x
66                    End If
                      
67                Next j
              
68            End If
69        Next i
          
          ' Update last sync remote date with the last sync local date
70        qry = vbNullString
71        qry = qry & "UPDATE " & SYNC_DATE_TBL & " " & _
                      "SET LastSyncDate = #" & Format(dteLastSyncLocal, "dd-mmm-yyyy Hh:Nn:Ss AM/PM") & "# " & _
                      "WHERE ID = '" & GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "ID", "WinLoginName", UCase(Environ("UserName")), """") & "'"
72        cnRemote.Execute qry
          
73        MsgBox "Data sync completed.", vbInformation
          
      ' Have new updates in remote db. The Admin updates COOP records
74    ElseIf dteLastSyncLocal < dteLastSyncRemote Then
          ' Loop through COOP tables
75        For i = 0 To UBound(arrMainTbl)
                 
              ' Get only records that are later than the last sync local date
76            qry = vbNullString
77            qry = qry & "SELECT T1.* " & _
                          "FROM " & arrMainTbl(i) & " AS T1 INNER JOIN " & PRA_EMPLOYEE_TBL & " AS T2 ON T1.CreatorID = T2.ID " & _
                          "WHERE T1.[LastSyncDate] > #" & dteLastSyncLocal & "# "
              ' Get every record that are owned by the user
              ' else get every later records if user has admin permission
78            If strUserPermission = enumUserPermission.OrdinaryUser Then
79                qry = qry & "AND T2.ID = '" & strUserID & "'"
80            End If
              
81            arrMainData = GetArrayList(qry, True, False)
              
82            If IsArrayAllocated(arrMainData) Then
                  ' Loop through all the remote data
83                For j = 0 To UBound(arrMainData)
                      ' Flag to update the other tables
84                    blnToUpdate = False
                      
                      ' Check if Ref# is already existing.
                      ' If yes update, else add a new record
85                    If IsItemExistInTable(CStr(arrMainTbl(i)), "RefNumber", CStr(arrMainData(j, 0)), "'", True) Then
                          ' Update
                          ' Check if the remote sync date is later than the local sync date
86                        If CDate(arrMainData(j, UBound(arrMainData, 2))) > _
                             CDate(GetItemFromMappingTbl(CStr(arrMainTbl(i)), "LastSyncDate", "RefNumber", CStr(arrMainData(j, 0)), "'", True)) Then
                              
                              ' Open recordset to enable changes
87                            qry = "SELECT * FROM " & arrMainTbl(i) & " " & _
                                    "WHERE RefNumber='" & arrMainData(j, 0) & "'"
88                            rs.Open qry, cn, adOpenKeyset, adLockOptimistic, adCmdText
                              
                              ' Change data in edit buffer
                              ' Do not update Ref#, Creator, Authoriser
89                            For x = 3 To UBound(arrMainData, 2)
90                                rs.Fields(x).Value = arrMainData(j, x)
91                            Next x
92                            rs.update
93                            Call CloseRecordset(rs)
                              
                              ' When successfully updated
94                            blnToUpdate = True
95                        Else    ' No updates done
                              ' Store table, ref# to the not updated Array
96                            DoEvents
97                        End If
98                    Else
                          ' Append
                          ' Open table to enable appending
99                        rs.Open arrMainTbl(i), cn, adOpenKeyset, adLockOptimistic, adCmdTable
                          
                          ' Add new record
100                       rs.AddNew
101                       For x = 0 To UBound(arrMainData, 2)
102                           rs.Fields(x).Value = arrMainData(j, x)
103                       Next x
104                       rs.update
105                       Call CloseRecordset(rs)
                          
106                       blnToUpdate = True
107                   End If
                      
                      ' Update other tables if the update or append are successfull
108                   If blnToUpdate Then
                          '
109                       arrOtherTbl = arrOpOtherTbl
                          
110                       For x = 0 To UBound(arrOtherTbl)
                              ' Get other local table records with the Ref#
111                           qry = vbNullString
112                           qry = qry & "SELECT T1.* " & _
                                          "FROM " & arrOtherTbl(x) & " AS T1 " & _
                                          "WHERE T1.RefNumber = '" & CStr(arrMainData(j, 0)) & "' "
113                           arrOtherLocal = GetArrayList(qry, True, False)
                              
114                           If IsArrayAllocated(arrOtherLocal) Then
                              
                                  ' Delete from remote
115                               qry = vbNullString
116                               qry = "DELETE * " & _
                                        "FROM " & arrOtherTbl(x) & " " & _
                                        "WHERE RefNumber = '" & CStr(arrMainData(j, 0)) & "';"
117                               cn.Execute qry
                          
                                  ' Append to local
                                  ' Open table to enable appending
118                               rs.Open arrOtherTbl(x), cn, adOpenKeyset, adLockOptimistic, adCmdTable
                                  
                                  ' Add new record
119                               For z = 0 To UBound(arrOtherLocal)
120                                   rs.AddNew
                                      
121                                   For y = 0 To UBound(arrOtherLocal, 2)
122                                       rs.Fields(y).Value = arrOtherLocal(z, y)
123                                   Next y
                                      
124                                   rs.update
125                               Next z
                                  
126                               Call CloseRecordset(rs)
                              
127                           End If
                              
128                       Next x
129                   End If
                      
130               Next j
              
131           End If
132       Next i
          
          ' Update last sync local date with the last sync remote date
133       qry = vbNullString
134       qry = qry & "UPDATE " & SYNC_DATE_TBL & " " & _
                      "SET LastSyncDate = #" & Format(dteLastSyncRemote, "dd-mmm-yyyy Hh:Nn:Ss AM/PM") & "# " & _
                      "WHERE ID = '" & GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "ID", "WinLoginName", UCase(Environ("UserName")), """", False) & "'"
135       cn.Execute qry
          
136       MsgBox "Data sync completed.", vbInformation
          
137   ElseIf dteLastSyncLocal = dteLastSyncRemote Then
138       MsgBox "Data already in sync.", vbExclamation
139   End If


      ' Update static tables
140   arrStaticTbl = Array(CUSTOMER_MAP_TBL, OUTLET_MAP_TBL, WHOLESALER_MAP_TBL, _
                           PRODUCT_MAP_TBL, PRICING_MAP_TBL, EXCISE_MAP_TBL, COGSPERLTR_MAP_TBL, KWI_MAP_TBL, PRA_EMPLOYEE_TBL, PRA_MANAGER_TBL)
141   strPWD = Chr(112) & Chr(114) & Chr(97) & Chr(117) & Chr(36)

142   If MsgBox("Do you want to update the Static tables?", vbYesNo, "Static Table Update") = vbYes Then
143       Application.Cursor = xlWait
144       For x = 0 To UBound(arrStaticTbl)
              ' Delete local table
145           qry = "DROP TABLE " & arrStaticTbl(x)
146           cn.Execute qry

              ' Copy table from remote db
147           Call CopyTableFromDB(REMOTE_DB_LOCATION & DB_NAME, CStr(arrStaticTbl(x)), strPWD, _
                                   ThisWorkbook.Path & "\" & DB_NAME, CStr(arrStaticTbl(x)), strPWD)
148       Next x
149       Application.Cursor = xlDefault
          
150       MsgBox "Static tables updated.", vbInformation, "Update Complete"
151   End If


      ' Disconnect from remote db
152   Call CloseDBConnection(cnRemote)


Proc_Exit:
153   PopCallStack
154   Exit Sub

Err_Handler:
155   GlobalErrHandler
156   Resume Proc_Exit
End Sub

Public Function GetColumnNames(strTblName As String) As Variant
      Dim rs As ADODB.Recordset
      Dim i As Integer
      Dim arr As Variant

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mDataSync|GetColumnNames"

3     Set rs = New ADODB.Recordset
4     rs.Open "SELECT TOP 1 * FROM " & strTblName, cn

5     ReDim arr(0)
6     For i = 0 To rs.Fields.Count - 1
7         If i <> 0 Then _
              ReDim Preserve arr(UBound(arr) + 1)
8         arr(i) = rs.Fields(i).Name
9     Next i

10    GetColumnNames = arr

Proc_Exit:
11    PopCallStack
12    Exit Function

Err_Handler:
13    GlobalErrHandler
14    Resume Proc_Exit
End Function

Public Sub CopyTableFromDB(strSourceDBPath As String, strSourceTbl As String, strSourcePwd As String, _
                           strDestinationDBPath As String, strDestinationTbl As String, strDestinationPwd As String)

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mDataSync|CopyTableFromDB"

3     cn.Execute "SELECT * INTO [MS Access;PWD=" & strDestinationPwd & ";DATABASE=" & strDestinationDBPath & "].[" & strDestinationTbl & "] " & _
                 "FROM [MS Access;PWD=" & strSourcePwd & ";DATABASE=" & strSourceDBPath & "].[" & strSourceTbl & "];"

Proc_Exit:
4     PopCallStack
5     Exit Sub

Err_Handler:
6     GlobalErrHandler
7     Resume Proc_Exit

End Sub
