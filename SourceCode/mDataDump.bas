Attribute VB_Name = "mDataDump"
Option Explicit

Const PERIOD_TEMP_TBL = "T_Temp_Period"
Const OUTLET_TEMP_TBL = "T_Temp_Outlet"
Const PROD_TEMP_TBL = "T_Temp_Prod"

Public Sub CreateDataDumpReport(wb As Workbook)
Dim ws As Worksheet
Dim rs As ADODB.Recordset
Dim strRefNumber As String
Dim i As Integer, x As Integer
Dim arrMain As Variant
Dim arrContractMonths As Variant
Dim intContractDuration As Integer
Dim lngOutletCount As Long
Dim lngFastarCount As Long
Dim lngFastarMult As Long
Dim strCopyFields As String
Dim qry As String

Const CELL_START = "A2"
Const FLD_REF_NUM = 0
Const FLD_FROM_DATE = 17
Const FLD_TO_DATE = 18
Const FLD_CONTRACT_LVL = 6
Const FLD_CONTRACT_LVL_CODE = 7

If gEnableErrorHandling Then On Error GoTo Err_Handler
PushCallStack "mDataDump|CreateDataDumpReport"

wb.Application.DisplayAlerts = False

' Copy sheet template
wb.Worksheets(DATA_DUMP_SHEET).Copy After:=wb.Worksheets(wb.Worksheets.Count)
Set ws = wb.ActiveSheet

' Rename sheet
ws.Name = DATA_DUMP_SHEET_RENAME

If LinkTransDB = True Then
    
    Set rs = New ADODB.Recordset
    
    ' Get Main table info
    qry = "SELECT * FROM " & OP_MAIN_TBL
    rs.Open qry, cn
    If Not rs.EOF Then
        arrMain = rs.GetRows(10000)
    End If
    Call CloseRecordset(rs)

    For x = 0 To UBound(arrMain, 2)

        strRefNumber = arrMain(FLD_REF_NUM, x)
    
        ' Get affected months
        arrContractMonths = GetAffectedMonths(CDate(arrMain(FLD_FROM_DATE, x)), CDate(arrMain(FLD_TO_DATE, x)))
        intContractDuration = UBound(arrContractMonths)
        cn.Execute "DELETE * FROM " & PERIOD_TEMP_TBL
        For i = 0 To intContractDuration - 1
            cn.Execute "INSERT INTO " & PERIOD_TEMP_TBL & "(RefNumber, Period) VALUES ('" & strRefNumber & "', '" & arrContractMonths(i) & "')"
        Next i
        
        ' Get list of Outlet info
        cn.Execute "DROP TABLE " & OUTLET_TEMP_TBL
        qry = "SELECT * INTO " & OUTLET_TEMP_TBL & " FROM (" & _
              "SELECT DISTINCT '" & strRefNumber & "' AS RefNumber, MatchCode, ExternalID, OutletName, State, BannerRegionCode " & _
              "FROM " & CUSTOMER_MAP_TBL & " " & _
              "WHERE "
        Select Case arrMain(FLD_CONTRACT_LVL, x)
            Case "OP Banner"
                qry = qry & "BannerCode = '" & arrMain(FLD_CONTRACT_LVL_CODE, x) & "')"
            Case "OP Banner Region"
                qry = qry & "BannerRegionCode = '" & arrMain(FLD_CONTRACT_LVL_CODE, x) & "')"
            Case "OP Outlet Level"
                qry = qry & "ExternalID IN (" & GetIN_List(Split(arrMain(FLD_CONTRACT_LVL_CODE, x), "|"), "'") & "))"
        End Select
        cn.Execute qry
        lngOutletCount = RecCount(OUTLET_TEMP_TBL)
        
        ' Set Fastar multiplier
        lngFastarMult = intContractDuration * lngOutletCount
        
        ' Get list of Product info
        cn.Execute "DROP TABLE " & PROD_TEMP_TBL
        qry = "SELECT * INTO " & PROD_TEMP_TBL & " FROM (" & _
              "SELECT DISTINCT T1.RefNumber, T1.ProductCode, T1.SubBrandCode, T2.SUB_BRAND_NAME, T2.FAMILY_NAME, T1.ProductType, T2.CATEGORY_NAME " & _
              "FROM " & OP_PROD_DETAILS_TBL & " AS T1 INNER JOIN " & PRODUCT_MAP_TBL & " AS T2 ON T1.SubBrandCode = T2.SUB_BRAND_CODE " & _
              "WHERE T1.RefNumber = '" & strRefNumber & "')" 'AND T1.StatusID = 3)"
        cn.Execute qry
        lngFastarCount = RecCount(PROD_TEMP_TBL)
        
        ' Generate Contract data
        qry = "SELECT 'Contract', T8.Description, T1.RefNumber, T1.FromDate, T1.ToDate, T1.FromDate_Extention, T1.ToDate_Extention, DateDiff('m',T1.FromDate,DateAdd('d',1,T1.ToDate)) AS Duration, " & _
              "T1.ContractType, T1.RouteToMarket, T2.Name, T1.ContractLevel, T1.OutletOrGroupName, T3.MatchCode, T3.ExternalID, T3.OutletName, T3.State, T3.BannerRegionCode, T4.SubBrandCode, " & _
              "T4.SUB_BRAND_NAME, T4.FAMILY_NAME, T4.ProductType, T4.CATEGORY_NAME, T7.Period, T5.ContractedVolume/" & lngFastarMult & " AS Ltr, T5.ContractedGSV/" & lngFastarMult & " AS GSV, T5.KWI/" & lngFastarMult & " AS KWI, " & _
              "T6.BannerTerms/" & lngFastarMult & " AS BannTerms, T6.StandardTerms/" & lngFastarMult & " AS StandTerms, T6.AdditionalTerms/" & lngFastarMult & " AS CondTerms, T5.COP/" & lngFastarMult & " AS COP, T5.QA3/" & lngFastarMult & " AS QA3, '', '', " & _
              "T5.COOP/" & lngFastarMult & " AS COOP, (T5.KWI/" & lngFastarMult & ")+(T6.BannerTerms/" & lngFastarMult & ")+(T6.StandardTerms/" & lngFastarMult & ")+(T6.AdditionalTerms/" & lngFastarMult & ")+(T5.COP/" & lngFastarMult & ")+(T5.QA3/" & lngFastarMult & ")+(T5.COOP/" & lngFastarMult & ") AS [AnD], " & _
              "(T5.ContractedGSV/" & lngFastarMult & ")-((T5.KWI/" & lngFastarMult & ")+(T6.BannerTerms/" & lngFastarMult & ")+(T6.StandardTerms/" & lngFastarMult & ")+(T6.AdditionalTerms/" & lngFastarMult & ")+(T5.COP/" & lngFastarMult & ")+(T5.QA3/" & lngFastarMult & ")+(T5.COOP/" & lngFastarMult & ")) AS NSV, " & _
              "T5.COGSnDistr/" & lngFastarMult & " AS COGSnDistr, (T5.ContractedGSV/" & lngFastarMult & ")-((T5.KWI/" & lngFastarMult & ")+(T6.BannerTerms/" & lngFastarMult & ")+(T6.StandardTerms/" & lngFastarMult & ")+(T6.AdditionalTerms/" & lngFastarMult & ")+(T5.COP/" & lngFastarMult & ")+(T5.QA3/" & lngFastarMult & ")+(T5.COOP/" & lngFastarMult & "))-T5.COGSnDistr/" & lngFastarMult & " AS CM, " & _
              "T5.AnP/" & lngFastarMult & " AS AnP, (T5.ContractedGSV/" & lngFastarMult & ")-((T5.KWI/" & lngFastarMult & ")+(T6.BannerTerms/" & lngFastarMult & ")+(T6.StandardTerms/" & lngFastarMult & ")+(T6.AdditionalTerms/" & lngFastarMult & ")+(T5.COP/" & lngFastarMult & ")+(T5.QA3/" & lngFastarMult & ")+(T5.COOP/" & lngFastarMult & "))-T5.COGSnDistr/" & lngFastarMult & "-T5.AnP/" & lngFastarMult & " AS CAAP, T1.PROS " & _
              "FROM ((((((" & OP_MAIN_TBL & " AS T1 INNER JOIN " & PRA_EMPLOYEE_TBL & " AS T2 ON T1.CreatorID = T2.ID) " & _
               "INNER JOIN " & OUTLET_TEMP_TBL & " AS T3 ON T1.RefNumber = T3.RefNumber) " & _
               "INNER JOIN " & PROD_TEMP_TBL & " AS T4 ON T1.RefNumber = T4.RefNumber) " & _
               "INNER JOIN " & OP_PROD_DETAILS_TBL & " AS T5 ON (T4.ProductCode = T5.ProductCode) AND (T4.RefNumber = T5.RefNumber)) " & _
               "INNER JOIN " & OP_TRADING_TERMS_TBL & " AS T6 ON (T4.ProductCode = T6.ProductCode) AND (T4.RefNumber = T6.RefNumber)) " & _
               "INNER JOIN " & PERIOD_TEMP_TBL & " AS T7 ON T1.RefNumber = T7.RefNumber) " & _
               "INNER JOIN " & STATUS_TBL & " AS T8 ON T1.StatusID = T8.ID " & _
              "WHERE T1.RefNumber = '" & strRefNumber & "' AND T1.StatusID <> 5"
        rs.Open qry, cn
    
        If Not rs.EOF Then
            'ws.Range(CELL_START).CopyFromRecordset rs
            With ws
                .Range(GetLastRow(.Range(CELL_START))).Offset(1, 0).CopyFromRecordset rs
                
                ' Get repeating fields for Actuals
                rs.MoveFirst
                strCopyFields = ""
                For i = 1 To 12
                    strCopyFields = strCopyFields & "'" & rs.Fields(i).Value & "', "
                Next i
            End With
        End If
        Call CloseRecordset(rs)
        
        ' Generate Actuals data
        qry = "SELECT DISTINCT 'Actuals', " & strCopyFields & " " & _
              "T1.Match_Code, T2.ExternalID, T2.OutletName, T2.State, T1.Ban_Reg_Code, " & _
              "T1.Fastar, T3.SUB_BRAND_NAME, T3.FAMILY_NAME, '', T3.CATEGORY_NAME, " & _
              "T1.MonthDate, T1.Qty_Ltr, T1.GSV, T1.KWI, '' AS BannTerms, T1.TT, '' AS CondTerms, T1.COP_Terms, T1.QA3, '' AS Expr2, '' AS Expr3, T1.COOP, " & _
              "T1.KWI + T1.TT + T1.COP_Terms + T1.QA3 + T1.COOP AS [AnD], " & _
              "T1.NSV, T1.CoGS + T1.Distrib AS COGSnDistr, T1.Net_Contribution, '' AS AnP, '' AS CAAP, '' AS PROS " & _
              "FROM (" & TRANS_TBL & " AS T1 INNER JOIN " & OUTLET_TEMP_TBL & " AS T2 ON T1.Match_Code = T2.MatchCode) " & _
              "INNER JOIN " & PRODUCT_MAP_TBL & " AS T3 ON T1.Fastar = T3.SUB_BRAND_CODE"
        rs.Open qry, cn
        
        If Not rs.EOF Then
            With ws
                .Range(GetLastRow(.Range(CELL_START))).Offset(1, 0).CopyFromRecordset rs
            End With
        End If
        Call CloseRecordset(rs)
    
    Next x
    
End If

Call CloseRecordset(rs, True)

' Remove linked Transaction table
cn.Execute "DROP TABLE " & TRANS_TBL

Proc_Exit:
PopCallStack
Exit Sub

Err_Handler:
GlobalErrHandler
Resume Proc_Exit
End Sub

Private Function RecCount(tbl As String) As Long
      Dim rs As ADODB.Recordset

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mDataDump|RecCount"

3     RecCount = 0

4     Set rs = New ADODB.Recordset

5     rs.Open "SELECT COUNT(*) FROM " & tbl, cn

6     If Not rs.EOF Then
7         RecCount = SetEmptyValue(rs.Fields(0), ZeroValue)
8     End If

9     Call CloseRecordset(rs, True)

Proc_Exit:
10    PopCallStack
11    Exit Function

Err_Handler:
12    GlobalErrHandler
13    Resume Proc_Exit
End Function

Private Function LinkTransDB() As Boolean
      Dim strFile As String
      Dim i As Integer
      Dim tbl As TableDef
      Dim arrTable As Variant

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mDataDump|LinkTransDB"

3     LinkTransDB = False

4     arrTable = Array(TRANS_TBL)

5     MsgBox "Please provide the location of the ""Main Transactions.accdb"" file." & vbCrLf & vbCrLf & "Press OK to continue.", vbInformation

      ' Select file
6     strFile = File_Dialog(msoFileDialogFilePicker, "Access DB Files", "*.accdb")

      ' Link Table
7     If Len(strFile) <> 0 Then
8         For i = 0 To UBound(arrTable)
9             Call CreateLinkedAccessTable(strFile, CStr(arrTable(i)), CStr(arrTable(i)))
10        Next i

11        LinkTransDB = True
12    End If

Proc_Exit:
13    PopCallStack
14    Exit Function

Err_Handler:
15    GlobalErrHandler
16    Resume Proc_Exit
          
End Function

Private Sub CreateLinkedAccessTable(strDBLinkTo As String, strLinkTbl As String, strLinkTblAs As String)
      Dim catDB As ADOX.Catalog
      Dim db As ADODB.Connection
      Dim tblLink As ADOX.Table
      Dim tbl As ADOX.Table

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mDataDump|CreateLinkedAccessTable"

3     Set catDB = New ADOX.Catalog

      ' Open a Catalog on the database in which to create the link.
4     catDB.ActiveConnection = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                               "Data Source=" & ThisWorkbook.Path & "\" & DB_NAME & ";" & _
                               "Jet OLEDB:Database " & Chr(80) & Chr(97) & Chr(115) & _
                               Chr(115) & Chr(119) & Chr(111) & Chr(114) & Chr(100) & _
                               Chr(61) & Chr(112) & Chr(114) & Chr(97) & Chr(117) & Chr(36) & ";"
          
      ' Delete linked table if already exist
5     For Each tbl In catDB.Tables
6         If tbl.Name = strLinkTblAs Then
7             cn.Execute "DROP TABLE " & strLinkTblAs
8             Exit For
9         End If
10    Next tbl
          
11    Set tblLink = New ADOX.Table
12    With tblLink
          ' Name the new Table and set its ParentCatalog property to the
          ' open Catalog to allow access to the Properties collection.
13        .Name = strLinkTblAs
14        Set .ParentCatalog = catDB
          
          ' Set the properties to create the link.
15        .Properties("Jet OLEDB:Create Link") = True
16        .Properties("Jet OLEDB:Link Datasource") = strDBLinkTo
17        .Properties("Jet OLEDB:Remote Table Name") = strLinkTbl
18    End With

      ' Append the table to the Tables collection.
19    catDB.Tables.Append tblLink

20    catDB.ActiveConnection.Close
21    Set catDB = Nothing

Proc_Exit:
22    PopCallStack
23    Exit Sub

Err_Handler:
24    GlobalErrHandler
25    Resume Proc_Exit
End Sub

Public Function GetAffectedMonths(dteFromDate As Date, dteToDate As Date) As Variant
      Dim strTemp As String
      Dim arrMonth As Variant
      Dim i As Long

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mDataDump|GetAffectedMonths"

3     ReDim arrMonth(0)
4     strTemp = Format(dteFromDate, "yyyymm")
5     arrMonth(0) = strTemp
6     For i = 1 To DateDiff("m", dteFromDate, DateAdd("d", 1, dteToDate))
7         strTemp = Format(DateAdd("m", 1, DateSerial(CInt(Left(strTemp, 4)), CInt(Right(strTemp, 2)), 1)), "yyyymm")
8         ReDim Preserve arrMonth(UBound(arrMonth) + 1)
9         arrMonth(i) = strTemp
10    Next i

11    GetAffectedMonths = arrMonth

Proc_Exit:
12    PopCallStack
13    Exit Function

Err_Handler:
14    GlobalErrHandler
15    Resume Proc_Exit
End Function

