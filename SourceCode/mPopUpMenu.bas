Attribute VB_Name = "mPopUpMenu"
Option Explicit

Dim arrList As Variant

Public Const Mname As String = "MyPopUpMenu"

Public Sub CreateFilterPopUpMenu(lst As MSForms.ListBox, intCol As Integer)
      Dim MenuItem As CommandBarPopup
      Dim arrFilter As Variant
      Dim blnExist As Boolean
      Dim x As Integer, y As Integer, z As Integer

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mPopUpMenu|CreateFilterPopUpMenu"

      'Delete PopUp menu if it exist
3     Call DeletePopUpMenu

4     If lst.ListCount > 0 Then
          ' Store list data in array
5         ReDim arrList(0 To lst.ListCount - 1, 0 To lst.ColumnCount - 1)
6         ReDim arrFilter(0)
7         For x = 0 To UBound(arrList)
8             For y = 0 To UBound(arrList, 2)
9                 arrList(x, y) = lst.List(x, y)
                  
                  ' Get list of filters
10                If y = intCol Then
11                    If Not IsInArray(arrFilter, arrList(x, y)) Then
12                        If Len(arrFilter(0)) <> 0 Then
13                            ReDim Preserve arrFilter(UBound(arrFilter) + 1)
14                        End If
15                        arrFilter(UBound(arrFilter)) = arrList(x, y)
16                    End If
17                End If
18            Next y
19        Next x
          
          'Create the PopUp menu
20        Select Case lst.Tag
              Case "Unfiltered"
21                With Application.CommandBars.Add(Name:=Mname, Position:=msoBarPopup, MenuBar:=False, Temporary:=True)
22                    For x = 0 To UBound(arrFilter)
23                        With .Controls.Add(Type:=msoControlButton)
24                            .Caption = arrFilter(x)
25                            .FaceId = 601
26                            .OnAction = "'PopulateFilteredList g_frmMain." & lst.Name & "," & intCol & ", """ & arrFilter(x) & """" & "'"
27                        End With
28                    Next x
29                End With
                  
30            Case "Filtered"
31                With Application.CommandBars.Add(Name:=Mname, Position:=msoBarPopup, MenuBar:=False, Temporary:=True)
32                    With .Controls.Add(Type:=msoControlButton)
33                        .Caption = "Remove Filter"
34                        .FaceId = 605
35                        .OnAction = "'PopulateFrontPageList g_frmMain." & lst.Name & "'"
36                    End With
                      
37                    For x = 0 To UBound(arrFilter)
38                        With .Controls.Add(Type:=msoControlButton)
39                            .Caption = arrFilter(x)
40                            .FaceId = 601
41                            .OnAction = "'PopulateFilteredList g_frmMain." & lst.Name & "," & intCol & ", """ & arrFilter(x) & """" & "'"
42                        End With
43                    Next x
44                End With
45        End Select
          
          ' Show the PopUp menu
46        Call ShowPopUpMenu
47    End If

Proc_Exit:
48    PopCallStack
49    Exit Sub

Err_Handler:
50    GlobalErrHandler
51    Resume Proc_Exit

End Sub

Public Sub PopulateFilteredList(lst As MSForms.ListBox, intCol As Integer, strFilter As String)
      Dim x As Integer, y As Integer

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mPopUpMenu|PopulateFilteredList"

      ' Clear listbox
3     lst.Clear

      ' Populate listbox with filtered data
4     For x = 0 To UBound(arrList)
          ' Check if filter is in the array
5         With lst
6             If arrList(x, intCol) = strFilter Then
7                 .AddItem
8                 For y = 0 To UBound(arrList, 2)
9                     .List(.ListCount - 1, y) = arrList(x, y)
10                Next y
11            End If
12        End With
13    Next x

14    If strFilter = vbNullString Then
15        lst.Tag = "Unfiltered"
16    Else
17        lst.Tag = "Filtered"
18    End If

Proc_Exit:
19    PopCallStack
20    Exit Sub

Err_Handler:
21    GlobalErrHandler
22    Resume Proc_Exit
End Sub

Public Sub DeletePopUpMenu()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mPopUpMenu|DeletePopUpMenu"
          
      'Delete PopUp menu if it exist
3     On Error Resume Next
4     Application.CommandBars(Mname).Delete
5     On Error GoTo 0

Proc_Exit:
6     PopCallStack
7     Exit Sub

Err_Handler:
8     GlobalErrHandler
9     Resume Proc_Exit
End Sub

Public Sub ShowPopUpMenu()
1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mPopUpMenu|ShowPopUpMenu"
          
      'Show the PopUp menu
3     On Error Resume Next
4     Application.CommandBars(Mname).ShowPopup
5     On Error GoTo 0

Proc_Exit:
6     PopCallStack
7     Exit Sub

Err_Handler:
8     GlobalErrHandler
9     Resume Proc_Exit
End Sub

Public Sub PopulateFrontPageList(objListBox As MSForms.ListBox, Optional strCreator As String = vbNullString)
      Dim qry As String
      Dim arr As Variant
      Dim x As Integer, y As Integer

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mPopUpMenu|PopulateFrontPageList"

3     qry = "SELECT DISTINCT T1.RefNumber, " & _
             "T3.[Name], " & _
             "T5.[Name], " & _
             "IIF(INSTR(1, T1.ContractLevelCode, '|')<>0, 'Multiple Outlets', T1.ContractLevelCode), " & _
             "T1.OutletOrGroupName, " & _
             "Format(T1.FromDate,'dd-mmm-yyyy'), " & _
             "Format(T1.ToDate,'dd-mmm-yyyy'), " & _
             "Format(T1.SubmitDate,'dd-mmm-yyyy'), " & _
             "T2.Description " & _
            "FROM ((((" & OP_MAIN_TBL & " AS T1 " & _
              "LEFT JOIN " & STATUS_TBL & " AS T2 ON T1.StatusID = T2.ID) " & _
              "LEFT JOIN " & PRA_EMPLOYEE_TBL & " AS T3 ON T1.CreatorID = T3.ID) " & _
              "LEFT JOIN " & PRA_MANAGER_TBL & " AS T4 ON T3.ManagerID = T4.ID) " & _
              "LEFT JOIN " & PRA_EMPLOYEE_TBL & " AS T5 ON T4.Name = T5.ID) " & _
            "WHERE "
                     
      ' Access for Admin
4     If g_iAccessType = enumUserPermission.Admin Then
5         qry = qry & "T1.StatusID <> " & enumStatus.statDeleted & " "
6     End If
                     
      ' Access for Managers
7     If g_iAccessType = enumUserPermission.Manager Then
8         qry = qry & "(T1.StatusID <> " & enumStatus.statDeleted & " AND T3.WinLoginName = """ & g_sLoginID & """) " & _
                   "OR (T1.StatusID IN (" & enumStatus.statForApproval & ", " & enumStatus.statApproved & ") AND T4.NAME = """ & GetItemFromMappingTbl(PRA_EMPLOYEE_TBL, "ID", "WinLoginName", g_sLoginID, """") & """)"
9     End If
                     
      ' Access for Ordinary users
10    If g_iAccessType = enumUserPermission.OrdinaryUser Then
11        qry = qry & "T1.StatusID <> " & enumStatus.statDeleted & " " & _
                      "AND T3.WinLoginName = """ & g_sLoginID & """"
12    End If

13    arr = GetArrayList(qry, True)

14    With objListBox
15        .Clear
16        If IsArrayAllocated(arr) Then
17            For x = 0 To UBound(arr)
18                .AddItem
19                For y = 0 To UBound(arr, 2)
20                    .List(x, y) = arr(x, y)
21                Next y
22            Next x
23        End If
24    End With

25    If strCreator = vbNullString Then
26        objListBox.Tag = "Unfiltered"
27    Else
28        objListBox.Tag = "Filtered"
29    End If

Proc_Exit:
30    PopCallStack
31    Exit Sub

Err_Handler:
32    GlobalErrHandler
33    Resume Proc_Exit
End Sub

