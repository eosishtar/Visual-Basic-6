Attribute VB_Name = "modCustReport"
Option Explicit

'global fields to exclude tables / and fields
Global EXCLUDE_TABLES As String
Global EXCLUDE_FIELDS As String
Global Const MAX_REPORTFIELDS = 40


Dim xlApp As Object
Dim xlBook As Object
Dim xlSheet As Object
Dim ReportHead As String


' -------------------------       TEMP STUFF   ---------------------------------------

Public cnDB As New ADODB.Connection
Public rsTables As ADODB.Recordset
Public rsFields As ADODB.Recordset
Public tField As ADODB.Field
Public TempRs As ADODB.Recordset

Public vPassword As String
Public daDb As String
Public cmd As String
Public sql As String

' -------------------------       TEMP STUFF   ---------------------------------------



Private Type ReportFields
  TableName As String
  FieldName As String
  DateType As Single
End Type

Public Type CustomReport
  ID As Integer
  RepName As String
  FieldCount As Integer
  DateFilter As Boolean
  OrderByField As String
  DataFields(1 To MAX_REPORTFIELDS) As ReportFields
End Type

Global Rep As CustomReport


Public Function LoadReportList(lv As ListBox)
   
  Dim XX As Integer
   
  lv.Clear
  Set TempRs = New ADODB.Recordset
    sql = "select ID,Desc from [CustReports]"
    TempRs.Open sql, cnDB, adOpenForwardOnly, adLockOptimistic
    
    lv.AddItem " -- Add New Report -- "
    lv.AddItem String(50, "-")
    lv.AddItem ""
    
    If TempRs.EOF Then
      TempRs.Close
      Exit Function
    End If
        
    XX = 4
    Do While Not TempRs.EOF
      lv.AddItem Trim$(TempRs!Desc) & "                                     " & TempRs!ID
      TempRs.MoveNext
    Loop
    
  TempRs.Close
  Set TempRs = Nothing


End Function

Public Sub LoadReportMainMenu(Optional ReloadMenu As Boolean)
   
  Dim XX As Integer
  Dim txt As String
  Dim ctl As Menu

    If ReloadMenu Then
      'only remove from 4 onwards
      On Error Resume Next
      For XX = 4 To frmCustReports.mnuSubReport.Count
        Unload frmCustReports.mnuSubReport(XX - 1)
      Next XX
      On Error GoTo 0
    End If
  
  Set TempRs = New ADODB.Recordset
    sql = "select ID,Desc from [CustReports]"
    TempRs.Open sql, cnDB, adOpenForwardOnly, adLockOptimistic

    
    If TempRs.EOF Then
      TempRs.Close
      Exit Sub
    End If
        
    XX = 4
    Do While Not TempRs.EOF
        Load frmCustReports.mnuSubReport(XX)
        frmCustReports.mnuSubReport(XX).Caption = Trim$(TempRs!Desc)
        frmCustReports.mnuSubReport(XX).Tag = TempRs!ID
        XX = XX + 1
      TempRs.MoveNext
    Loop
    
  TempRs.Close
  Set TempRs = Nothing

End Sub

Public Function GetReportSetting(ByVal vID As Integer)

  Dim RepCnt As Integer
  Dim tFields() As String
  Dim tTemp As String
  Dim Items() As String

  Rep.ID = 0
  Rep.RepName = ""
  Rep.FieldCount = 0
  Rep.DateFilter = False
  Rep.OrderByField = ""
  Call ClearRepList
      
  Set TempRs = New ADODB.Recordset
    sql = "select * from [CustReports] WHERE ID = " & vID
      TempRs.Open sql, cnDB, adOpenForwardOnly, adLockOptimistic
    
      If TempRs.EOF Then
        TempRs.Close
      Else
        'should never be blank
        Rep.ID = vID
        Rep.RepName = Trim$(TempRs!Desc)

        tTemp = Left(TempRs!Record, InStr(TempRs!Record, "|") - 1)
        Items = Split(tTemp, ",")
        Rep.FieldCount = Items(0)         'Nr of columns
        Rep.DateFilter = CBool(Items(1))  'Date Filter
        Rep.OrderByField = Items(2)       'Order By Field
        
        tTemp = Right(TempRs!Record, Len(TempRs!Record) - InStrRev(TempRs!Record, "|"))
        Items() = Split(tTemp, ",")
        For RepCnt = 1 To Rep.FieldCount
          tFields() = Split(Items(RepCnt - 1), ".")
          Rep.DataFields(RepCnt).TableName = tFields(0)
          Rep.DataFields(RepCnt).FieldName = tFields(1)
          Rep.DataFields(RepCnt).DateType = tFields(2)
        Next RepCnt
        
       TempRs.Close
    End If

  Set TempRs = Nothing
  
End Function

Public Function MakeCustomSQL(ReportString As String) As String
  Dim tTemp As String
  Dim Items() As String
  Dim RepCnt As Integer
  Dim tSQL As String
  Dim tFields() As String
  Dim tTables As String
  Dim tColCount As Integer
  Dim tDateOn As Boolean
  Dim tOrderBy As String
  
  'Get the Settings for the query
  tTemp = Left(ReportString, InStr(ReportString, "|") - 1)
  Items = Split(tTemp, ",")
  tColCount = Items(0)       'Nr of columns
  tDateOn = CBool(Items(1))  'Date Filter
  tOrderBy = Items(2)        'Order By Field
  
  'build the fields selected
  tTables = ""
  tTemp = Right(ReportString, Len(ReportString) - InStrRev(ReportString, "|"))
  Items() = Split(tTemp, ",")
  For RepCnt = 1 To tColCount
    tFields() = Split(Items(RepCnt - 1), ".")
     tSQL = tSQL & tFields(0) & "." & tFields(1) & ", "
     tTables = tTables & tFields(0) & ","     'collect all the table names
  Next RepCnt
  tSQL = Left(tSQL, Len(tSQL) - 2)     'remove last comma off sql statement
  
  'get the chosen columns
  tTables = GetSelectedTables(tColCount, tTables)
  
  If tOrderBy = "" Then
    'no order by selected, use RS
    MakeCustomSQL = "Select " & tSQL & " FROM " & tTables
  Else
    'order by chosen table/field
    MakeCustomSQL = "Select " & tSQL & " FROM " & tTables & "ORDER BY " & tOrderBy
  End If

End Function

Public Function GetSelectedTables(vColCount As Integer, vTables As String) As String

  Dim dX As Integer
  Dim dV As Integer
  Dim dFound As Boolean
  Dim dItem As String
  
  Dim dSelFields() As String
  Dim TableArray() As String
  Dim TableCnt As Integer
  
  ReDim TableArray(1 To vColCount) 'prepare to get all the table names
  dSelFields() = Split(vTables, ",")
  TableCnt = 1
  
  For dX = 1 To vColCount
    dFound = False
      'search through all tables for chosen fields and remove duplicates
      For dV = 1 To UBound(TableArray)
        If UCase(Trim$(dSelFields(dX))) = UCase(Trim$(TableArray(dV))) Then
          dFound = True
          Exit For
        End If
      Next dV
      
      If Not dFound Then
        'can add it
        TableArray(TableCnt) = dSelFields(dX)
        TableCnt = TableCnt + 1
      End If
  Next dX
  
  GetSelectedTables = ""
  For dX = 1 To TableCnt - 1
    GetSelectedTables = GetSelectedTables & TableArray(dX) & ", "
  Next dX
  GetSelectedTables = Left(GetSelectedTables, Len(GetSelectedTables) - 2)     'remove last comma off sql statement
  
End Function

Private Sub ClearRepList()
  Dim r As Integer

  For r = 1 To MAX_REPORTFIELDS
    Rep.DataFields(r).TableName = ""
    Rep.DataFields(r).FieldName = ""
    Rep.DataFields(r).DateType = 0
  Next r
  
End Sub

Public Function GetDataType(vDataType As Integer) As String
  
  GetDataType = ""
  
  Select Case vDataType
    Case 11
      GetDataType = "Boolean"
    Case 6
      GetDataType = "Currency"
    Case 13
      GetDataType = "Unknown"
    Case 200, 201, 20
      GetDataType = "String"
    Case 202
      GetDataType = "Short Text"
    Case 203
      GetDataType = "Long Text"
    Case 3, 5
      GetDataType = "Number"
    Case 205
      GetDataType = "OLE Object"
    Case 0
      GetDataType = ""
  End Select

End Function



'
'
'
'     -----------------------                 EXCEL REPORTER              -------------------------
'
'
'



Public Sub ExcelReportDump(vSql As String)

  Dim Nextrow As Integer
  Dim cColumnLetter As String
  Dim GetFile As String
  Dim i, j, k As Integer
  Dim t As String
  Dim flds() As String
  Dim vGeneralValue As String
  Dim vProfit As Double
  Dim vGotTotal As Boolean
  Dim fldCunt As Integer
  Dim recCunt As Integer
  
  Set TempRs = New ADODB.Recordset
  
  With TempRs
  .Open vSql, cnDB, adOpenKeyset, adLockOptimistic
  
  If .EOF Then
    .Close
    MsgBox "There is no data to report!", vbInformation, "Report Manager... "
    Exit Sub
  End If

  recCunt = .RecordCount
  fldCunt = .Fields.Count
  
  On Error Resume Next
  Set xlApp = CreateObject("excel.application")
  Screen.MousePointer = vbHourglass
  If Err.Number <> 0 Then
    Err.Clear
      Screen.MousePointer = vbDefault
        MsgBox "Microsoft Excel could not be detected on this PC.", vbExclamation + vbOKCancel, "Excel Reporting..."
    Exit Sub
  End If
  
  Set xlBook = xlApp.Workbooks.Add
  Set xlSheet = xlBook.Worksheets("Sheet1")
  xlApp.Visible = True
  On Error GoTo 0
  
  '                 Generic for all reports
  ' -----------------------------------------------------
  '
  
  '...setting the width
  'SetColWidth        ' STILL TO DO
  
  '...setting the headings
  'xlSheet.Cells(2, 1).Value = Company.TradingName
  'xlSheet.Range("A1").Resize(2, 1).Font.Bold = True
  'xlSheet.Range("A1").Resize(2, 1).Font.Underline = True
  'Nextrow = 2
    
  '... seperate address
  't = Replace(Company.Address, vbCrLf, ",")
  'flds = Split(t, ",")
  'For i = 1 To UBound(flds)
  '  If flds(i) <> "" Then
  '    xlSheet.Cells(Nextrow + i, 1).Value = flds(i - 1)
  '  End If
  'Next i
  
  '..set next row nr from array and add date
  Nextrow = Nextrow + UBound(flds) + 1
  xlSheet.Cells(Nextrow, 3).Value = Rep.RepName
  xlSheet.Cells(Nextrow, TempRs.Fields.Count).Value = "Report Date: " & Format(Date, "dd MMMM yyyy")
  cColumnLetter = GetAlphaLetter(.Fields.Count)
  xlSheet.Range(cColumnLetter & Nextrow).HorizontalAlignment = -4152    'align right / xlRight
    
  '...get corrsponding alpha letter based on column count
  cColumnLetter = GetAlphaLetter(.Fields.Count + 1)
      
  '...set the logo        STILL TO DO
  'With xlSheet.Pictures.Insert(GetAppPath & "MyLogo.bmp")
  '  With .ShapeRange
  '    .LockAspectRatio = -1
  '    .Width = 50
  '    .Height = 50
  '  End With
  '    .Left = xlSheet.Range(cColumnLetter & "1").Left - .Width
  '    .Top = xlSheet.Range("A1").Top
  '    .Placement = 1
  '    .PrintObject = True
  'End With
  
  Nextrow = Nextrow + 2     'skip a line
  .MoveFirst
    
  '... count the columns and centre them in the loop
  For i = 1 To TempRs.Fields.Count
    xlSheet.Cells(Nextrow, i).Value = TempRs.Fields(i - 1).Name
    xlSheet.Cells(Nextrow, i).HorizontalAlignment = -4108
    xlSheet.Cells(Nextrow, i).VerticalAlignment = -4108
  Next i
  
  '...style the header
  cColumnLetter = GetAlphaLetter(TempRs.Fields.Count)
  cColumnLetter = cColumnLetter & Nextrow
  xlSheet.Range("A" & Nextrow & ":" & cColumnLetter).Interior.Color = &HC0C0C0
  With xlSheet.Range("A" & Nextrow & ":" & cColumnLetter).Borders
    .LineStyle = 1  'xlContinuous
    .Color = vbBlack
    .Weight = 2     'xlThin
  End With
  Nextrow = Nextrow + 2     'skip a line
  
                  
  '                   Report specific
  ' -----------------------------------------------------
  '
 
  TempRs.MoveFirst
  
      For i = 1 To recCunt
        For j = 1 To fldCunt
          Select Case Rep.DataFields(j).DateType
            Case 11                      'Boolean
              If UCase(TempRs.Fields(j - 1).Value) = "FALSE" Then
                xlSheet.Cells(Nextrow, j).Value = "No"
              Else
                xlSheet.Cells(Nextrow, j).Value = "Yes"
              End If
            Case 6                       'Currency
               xlSheet.Cells(Nextrow, j).Value = FormatNumber(TempRs.Fields(j - 1).Value, 2)
            Case 13                      'Unknown
            Case 200, 201, 20, 202, 203  'String
              If InStr(1, UCase(Rep.DataFields(j).FieldName), "CODE") > 0 Then
                xlSheet.Cells(Nextrow, j).Value = "'" & TempRs.Fields(j - 1).Value
              ElseIf InStr(1, UCase(Rep.DataFields(j).FieldName), "COST") > 0 Then
                xlSheet.Cells(Nextrow, j).Value = FormatNumber(TempRs.Fields(j - 1).Value, 2)
              ElseIf InStr(1, UCase(Rep.DataFields(j).FieldName), "DATE") > 0 Then
                xlSheet.Cells(Nextrow, j).Value = Format(TempRs.Fields(j - 1).Value, "dd MMMM yyyy")
              ElseIf InStr(1, UCase(Rep.DataFields(j).FieldName), "TOTAL") > 0 Then
                xlSheet.Cells(Nextrow, j).Value = FormatNumber(TempRs.Fields(j - 1).Value, 2)
              Else
                xlSheet.Cells(Nextrow, j).Value = TempRs.Fields(j - 1).Value
              End If
              
            Case 3, 5                    'Number
              If InStr(1, UCase(Rep.DataFields(j).FieldName), "ID") > 0 Then
                xlSheet.Cells(Nextrow, j).Value = TempRs.Fields(j - 1).Value
              ElseIf InStr(1, UCase(Rep.DataFields(j).FieldName), "TOTAL") > 0 Then
                xlSheet.Cells(Nextrow, j).Value = FormatNumber(TempRs.Fields(j - 1).Value, 2)
              Else
                xlSheet.Cells(Nextrow, j).Value = TempRs.Fields(j - 1).Value
              End If
            Case 205                     'OLE Object
            Case 0
            Case Else
            
              'xlSheet.Cells(Nextrow, j).Value = Format(TempRs.Fields(j - 1).Value, "dd MMMM yyyy")
              'xlSheet.Cells(Nextrow, j).Value = FormatNumber(TempRs.Fields(j - 1).Value, 2)
              'xlSheet.Cells(Nextrow, j).Value = "'" & TempRs.Fields(j - 1).Value
              'xlSheet.Cells(Nextrow, j).Value = TempRs.Fields(j - 1).Value
              
          End Select
          xlSheet.Cells(Nextrow, j).HorizontalAlignment = -4108
          xlSheet.Cells(Nextrow, j).VerticalAlignment = -4108
        Next j
        TempRs.MoveNext
        Nextrow = Nextrow + 1
      Next i
      
  'auto fit
  xlSheet.Range("A1").Resize(1, j - 1).EntireColumn.AutoFit
  Screen.MousePointer = vbNormal
  
End With

Set TempRs = Nothing

End Sub


Public Function GetAlphaLetter(cColumnNr As Integer) As String

  If cColumnNr = 1 Then GetAlphaLetter = "A"
  If cColumnNr = 2 Then GetAlphaLetter = "B"
  If cColumnNr = 3 Then GetAlphaLetter = "C"
  If cColumnNr = 4 Then GetAlphaLetter = "D"
  If cColumnNr = 5 Then GetAlphaLetter = "E"
  If cColumnNr = 6 Then GetAlphaLetter = "F"
  If cColumnNr = 7 Then GetAlphaLetter = "G"
  If cColumnNr = 8 Then GetAlphaLetter = "H"
  If cColumnNr = 9 Then GetAlphaLetter = "I"
  If cColumnNr = 10 Then GetAlphaLetter = "J"
  If cColumnNr = 11 Then GetAlphaLetter = "K"
  If cColumnNr = 12 Then GetAlphaLetter = "L"
  If cColumnNr = 13 Then GetAlphaLetter = "M"
  If cColumnNr = 14 Then GetAlphaLetter = "N"
  If cColumnNr = 15 Then GetAlphaLetter = "O"
  If cColumnNr = 16 Then GetAlphaLetter = "P"
  If cColumnNr = 17 Then GetAlphaLetter = "Q"
  If cColumnNr = 18 Then GetAlphaLetter = "R"
  If cColumnNr = 19 Then GetAlphaLetter = "S"
  If cColumnNr = 20 Then GetAlphaLetter = "T"
  If cColumnNr = 21 Then GetAlphaLetter = "U"
  If cColumnNr = 22 Then GetAlphaLetter = "V"
  If cColumnNr = 23 Then GetAlphaLetter = "W"
  If cColumnNr = 24 Then GetAlphaLetter = "X"
  If cColumnNr = 25 Then GetAlphaLetter = "Y"
  If cColumnNr = 26 Then GetAlphaLetter = "Z"

End Function

Public Function SetColWidth(Rep As Integer)
  '...for every character, the column width must be min 1.43
  ' 1  char = 1.43
  ' 5  char = 6
  ' 10 char = 11.71
  ' 15 char = 17.71
    
 Select Case Rep
  Case 1       'product details
     ReportHead = "Product Listing"
     xlSheet.Range("A1").ColumnWidth = 4
     xlSheet.Range("B1").ColumnWidth = 11
     xlSheet.Range("C1").ColumnWidth = 40
     xlSheet.Range("D1").ColumnWidth = 15
     xlSheet.Range("E1").ColumnWidth = 15
     xlSheet.Range("F1").ColumnWidth = 15
     xlSheet.Range("G1").ColumnWidth = 15

    Case 2    'User listing
     ReportHead = "User Listing"
     xlSheet.Range("A1").ColumnWidth = 6
     xlSheet.Range("B1").ColumnWidth = 15
     xlSheet.Range("C1").ColumnWidth = 25
     xlSheet.Range("D1").ColumnWidth = 18
     xlSheet.Range("E1").ColumnWidth = 18

    Case 3    'client listing
     ReportHead = "Client Listing"
     xlSheet.Range("A1").ColumnWidth = 6
     xlSheet.Range("B1").ColumnWidth = 18
     xlSheet.Range("C1").ColumnWidth = 16
     xlSheet.Range("D1").ColumnWidth = 20
     xlSheet.Range("E1").ColumnWidth = 30
     xlSheet.Range("F1").ColumnWidth = 16
     xlSheet.Range("G1").ColumnWidth = 16
     
    Case 4      'master vehicle listing
     ReportHead = "Vehicle Details"
     xlSheet.Range("A1").ColumnWidth = 4
     xlSheet.Range("B1").ColumnWidth = 11
     xlSheet.Range("C1").ColumnWidth = 15
     xlSheet.Range("D1").ColumnWidth = 15
     xlSheet.Range("E1").ColumnWidth = 4
     xlSheet.Range("F1").ColumnWidth = 9
     xlSheet.Range("G1").ColumnWidth = 4
     xlSheet.Range("H1").ColumnWidth = 6
     xlSheet.Range("I1").ColumnWidth = 11
     xlSheet.Range("J1").ColumnWidth = 17
     xlSheet.Range("K1").ColumnWidth = 6
     xlSheet.Range("L1").ColumnWidth = 22
     xlSheet.Range("M1").ColumnWidth = 9
     xlSheet.Range("N1").ColumnWidth = 9
  
    Case 10     ' export Category
      ReportHead = "Category Listing"
      xlSheet.Range("A1").ColumnWidth = 5
      xlSheet.Range("B1").ColumnWidth = 20
      xlSheet.Range("C1").ColumnWidth = 15
      xlSheet.Range("D1").ColumnWidth = 30
      xlSheet.Range("E1").ColumnWidth = 20
      xlSheet.Range("F1").ColumnWidth = 20
  
    Case 11     ' export postal codes
      ReportHead = "Postal Code Listing"
      xlSheet.Range("A1").ColumnWidth = 5
      xlSheet.Range("B1").ColumnWidth = 28
      xlSheet.Range("C1").ColumnWidth = 16
      xlSheet.Range("D1").ColumnWidth = 12
      xlSheet.Range("E1").ColumnWidth = 23
      xlSheet.Range("F1").ColumnWidth = 28
    
  End Select

End Function




  



  'Dim cColWidth As Integer
  '...count columns, and set widths for all of them
  'xlSheet.Range("A1").ColumnWidth = 15
  'xlSheet.Range("B1").ColumnWidth = 30
  'For i = 1 To TempExcelrs.Fields.Count
  'cColumnLetter = GetAlphaLetter(TempExcelrs.Fields.Count)
  '  cColWidth = SetColWidth(Len(TempExcelrs.Fields(i - 1).Name), TempExcelrs.Fields.Count)
  '  xlSheet.Range(cColumnLetter & "1").ColumnWidth = 22
  'Next i
  'Worksheets("Sheet1").Range("A1:E1").Columns.AutoFit
  

