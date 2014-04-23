Attribute VB_Name = "mTransferTables"
Option Explicit

Public Sub TransferMDMData()
      Dim arrTable As Variant
      Dim arrTableDescription As Variant
      Dim arrCustDatabaseFld As Variant
      Dim arrCustSourceFld As Variant
      Dim arrProdDatabaseFld As Variant
      Dim arrProdSourceFld As Variant
      Dim arrDatabaseFld As Variant
      Dim arrSourceFld As Variant
      Dim strFile As String
      Dim i As Long, x As Long, y As Long, z As Long
      Dim wb As Workbook
      Dim ws As Worksheet
      Dim rng As Range
      Dim strValues As String
      Dim qry As String
      Dim blnImported As Boolean

1     If gEnableErrorHandling Then On Error GoTo Err_Handler
2     PushCallStack "mTransferTables|TransferMDMData"

3     arrTable = Array(CUSTOMER_MAP_TBL, PRODUCT_MAP_TBL)
4     arrTableDescription = Array("MDM Customers", "MDM Products")

5     arrCustDatabaseFld = Array("Banner", "BannerCode", "BannerRegion", "BannerRegionCode", "OutletName", "MatchCode", "ExternalID", "State")
6     arrCustSourceFld = Array("BANNER_NAME", "BANNER_CODE", "BANNER_REGION_NAME", "BANNER_REGION_CODE", "OUTLET_NAME", "LIQUOR_LICENSE", "#EXTERNAL_ID", "STATE")

7     arrProdDatabaseFld = Array("FAMILY_CODE", "FAMILY_NAME", "CATEGORY_CODE", "CATEGORY_NAME", "SUB_CATEGORY_CODE", "SUB_CATEGORY_NAME", "BRAND_CODE", "BRAND_NAME", "SUB_BRAND_CODE", "SUB_BRAND_NAME", "PRODUCT_CODE", "PRODUCT_DESCRIPTION", "BOTTLE_SIZE", "UNITS_PER_CASE")
8     arrProdSourceFld = Array("#FAMILY_CODE", "FAMILY_NAME", "CATEGORY_CODE", "CATEGORY_NAME", "SUB_CATEGORY_CODE", "SUB_CATEGORY_NAME", "BRAND_CODE", "BRAND_NAME", "SUB_BRAND_CODE", "SUB_BRAND_NAME", "PRODUCT_CODE", "PRODUCT_DESCRIPTION", "BOTTLE_SIZE", "UNITS_PER_CASE")

9     arrDatabaseFld = Array(arrCustDatabaseFld, arrProdDatabaseFld)
10    arrSourceFld = Array(arrCustSourceFld, arrProdSourceFld)

11    blnImported = False
12    For i = 0 To UBound(arrTable)
13        If MsgBox("Do you want to import " & arrTableDescription(i) & " data?", vbYesNo, "Import MDM Data") = vbYes Then
              ' Select file
14            strFile = File_Dialog(msoFileDialogFilePicker, "CSV Files", "*.csv")
              
15            If strFile <> "" Then
16                blnImported = True
                  
                  ' Create a new workbook
17                Set wb = Application.Workbooks.Open(strFile)
18                Set ws = wb.Worksheets(1)
                  
                  ' Set reference cell
19                Set rng = ws.Range("A1")
                  
                  ' Delete data from table
20                cn.Execute "DELETE * FROM " & arrTable(i)
          
21                x = 1
22                Do While Len(rng.Offset(x, 0).Value) <> 0
                      ' Get values
23                    strValues = vbNullString
24                    For y = 0 To UBound(arrSourceFld(i))
                          ' Search for value field
25                        z = 0
26                        Do While Len(rng.Offset(0, z).Value) <> 0
27                            If arrSourceFld(i)(y) = rng.Offset(0, z).Value Then
28                                strValues = strValues & """" & Replace(rng.Offset(x, z).Value, """", """""") & ""","
29                                Exit Do
30                            End If
31                            z = z + 1
32                        Loop
33                    Next y
34                    strValues = Left(strValues, Len(strValues) - 1)
                  
                      ' Insert line to table
35                    qry = "INSERT INTO " & arrTable(i) & " " & _
                            "VALUES(" & strValues & ")"
36                    cn.Execute qry
          
37                    x = x + 1
38                Loop
                  
                  ' Close input file
39                wb.Close
40            End If
41        End If
42    Next i

43    If blnImported Then
44        MsgBox "Finished importing MDM data.", vbInformation, "Import MDM Data"
45    Else
46        MsgBox "No MDM data imported.", vbInformation, "Import MDM Data"
47    End If

Proc_Exit:
48    PopCallStack
49    Exit Sub

Err_Handler:
50    GlobalErrHandler
51    Resume Proc_Exit
End Sub

