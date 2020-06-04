Attribute VB_Name = "modExcel"
Option Explicit
Dim xlApp As Object
Dim xlBook As Object
Dim xlSheet As Object
Dim ReportHead As String
Dim TempExcelrs As ADODB.Recordset



Public Sub ExcelDump(vSql As String, vReport As Integer)

  Dim Nextrow As Integer
  Dim cColumnLetter As String
  Dim GetFile As String
  Dim i, j As Integer
  Dim t As String
  Dim flds() As String
  Dim vGeneralValue As String
  Dim vProfit As Double
  Dim vTotals(1 To 4) As Double
  Dim vGotTotal As Boolean
  Dim fldCunt As Integer
  Dim recCunt As Integer
  Dim CommentCounter As Integer
  Dim CommentAVG As Double
  
  Set TempExcelrs = New ADODB.Recordset
  
  With TempExcelrs
  .Open vSql, cn, adOpenKeyset, adLockOptimistic
  
  If .EOF Then
    .Close
    MsgBox "There is no data to report!", vbInformation, "Report Manager... "
    Exit Sub
  End If

  recCunt = TempExcelrs.RecordCount
  fldCunt = TempExcelrs.Fields.Count
  CommentCounter = 0
  vTotals(1) = 0
  vTotals(2) = 0
  vTotals(3) = 0
  vTotals(4) = 0
  
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
  SetColWidth Val(vReport)
  
  '...setting the headings
  xlSheet.Cells(2, 1).Value = Company.Name
  xlSheet.Range("A1").Resize(2, 1).Font.Bold = True
  xlSheet.Range("A1").Resize(2, 1).Font.Underline = True
  Nextrow = 2
    
  '... seperate address
  t = Replace(Company.Address, vbCrLf, ",")
  flds = Split(t, ",")
  For i = 1 To UBound(flds)
    If flds(i) <> "" Then
      xlSheet.Cells(Nextrow + i, 1).Value = flds(i - 1)
    End If
  Next i
  
  '..set next row nr from array and add date
  Nextrow = Nextrow + UBound(flds) + 1
  xlSheet.Cells(Nextrow, 3).Value = ReportHead
  xlSheet.Cells(Nextrow, TempExcelrs.Fields.Count).Value = "Report Date: " & Format(Date, "dd MMMM yyyy")
  cColumnLetter = GetAlphaLetter(TempExcelrs.Fields.Count)
  xlSheet.Range(cColumnLetter & Nextrow).HorizontalAlignment = -4152    'align right / xlRight
    
  '...get corrsponding alpha letter based on column count
  cColumnLetter = GetAlphaLetter(TempExcelrs.Fields.Count + 1)
      
  '...set the logo
  With xlSheet.Pictures.Insert(GetAppPath & "MyLogo.jpg")
    With .ShapeRange
      .LockAspectRatio = -1
      .Width = 50
      .Height = 50
    End With
      .Left = xlSheet.Range(cColumnLetter & "1").Left - .Width
      .Top = xlSheet.Range("A1").Top
      .Placement = 1
      .PrintObject = True
  End With
  
  Nextrow = Nextrow + 2     'skip a line
  TempExcelrs.MoveFirst
    
  '... count the columns and centre them in the loop
  For i = 1 To TempExcelrs.Fields.Count
    xlSheet.Cells(Nextrow, i).Value = TempExcelrs.Fields(i - 1).Name
    xlSheet.Cells(Nextrow, i).HorizontalAlignment = -4108
    xlSheet.Cells(Nextrow, i).VerticalAlignment = -4108
  Next i
    
    
  '...style the header
  cColumnLetter = GetAlphaLetter(TempExcelrs.Fields.Count)
  cColumnLetter = cColumnLetter & Nextrow
  xlSheet.Range("A" & Nextrow & ":" & cColumnLetter).Interior.Color = &HC0C0C0
  With xlSheet.Range("A" & Nextrow & ":" & cColumnLetter).Borders
    .LineStyle = 1 'xlContinuous
    .Color = vbBlack
    .Weight = 2 'xlThin
  End With
  Nextrow = Nextrow + 2     'skip a line
  
                  
  '                   Report specific
  ' -----------------------------------------------------
  '
 
  
  Select Case vReport
    Case 1   'single vehicle details
      TempExcelrs.MoveFirst
      For i = 1 To recCunt
        For j = 1 To fldCunt
          If j = 1 Then
            xlSheet.Cells(Nextrow, j).Value = "'" & TempExcelrs.Fields(j - 1).Value
          Else
            If j = 3 Then
              GetFile = Right(TempExcelrs.Fields(j - 1).Value, Len(TempExcelrs.Fields(j - 1).Value) - InStrRev(TempExcelrs.Fields(j - 1).Value, "\"))
              xlSheet.Cells(Nextrow, j).Value = GetFile
            Else
              xlSheet.Cells(Nextrow, j).Value = TempExcelrs.Fields(j - 1).Value
            End If
          End If
          xlSheet.Cells(Nextrow, j).HorizontalAlignment = -4108
        Next j
        TempExcelrs.MoveNext
        Nextrow = Nextrow + 1
      Next i
      
    Case 2   'Deal List
      TempExcelrs.MoveFirst
      For i = 1 To recCunt
        For j = 1 To fldCunt
          If j = 4 Or j = 7 Then    ' tel numbers
            xlSheet.Cells(Nextrow, j).Value = "'" & TempExcelrs.Fields(j - 1).Value
          Else
            If j = 8 Then       'get car details
              GetMasterVehicleDetails Val(TempExcelrs!VehicleID)
              xlSheet.Cells(Nextrow, j).Value = "(" & MasterVehicle.ModelYear & ")" & "   " & MasterVehicle.VehicleMake & " " & MasterVehicle.VehicleModel
            Else
              xlSheet.Cells(Nextrow, j).Value = TempExcelrs.Fields(j - 1).Value
            End If
          End If
          xlSheet.Cells(Nextrow, j).HorizontalAlignment = -4108
        Next j
        TempExcelrs.MoveNext
        Nextrow = Nextrow + 1
      Next i
      
    Case 3      'client listing
      TempExcelrs.MoveFirst
      For i = 1 To recCunt
        For j = 1 To fldCunt
          If j = 3 Or j = 5 Or j = 6 Or j = 7 Then
            xlSheet.Cells(Nextrow, j).Value = "'" & TempExcelrs.Fields(j - 1).Value
          ElseIf j = 10 Then
            xlSheet.Cells(Nextrow, j).Value = Format(TempExcelrs.Fields(j - 1).Value, "dd MMMM yyyy")
          Else
            xlSheet.Cells(Nextrow, j).Value = TempExcelrs.Fields(j - 1).Value
          End If
          xlSheet.Cells(Nextrow, j).HorizontalAlignment = -4108
        Next j
        TempExcelrs.MoveNext
        Nextrow = Nextrow + 1
      Next i
      
    Case 4      'Deal listing
      TempExcelrs.MoveFirst
      vGotTotal = False
      For i = 1 To recCunt
        For j = 1 To fldCunt
          'collect the totals
          If Not vGotTotal Then
            vProfit = TempExcelrs.Fields(7).Value - (TempExcelrs.Fields(5).Value + TempExcelrs.Fields(6).Value)
            vTotals(1) = vTotals(1) + TempExcelrs.Fields(7).Value   'Vehicle sold for
            vTotals(2) = vTotals(2) + TempExcelrs.Fields(5).Value   'Vehicle Purchase
            vTotals(3) = vTotals(3) + TempExcelrs.Fields(6).Value   'total spent on parts
            vTotals(4) = vTotals(4) + vProfit                       'total profit
            vGotTotal = True
          End If
          
          If j = 3 Then
            xlSheet.Cells(Nextrow, j).Value = "'" & TempExcelrs.Fields(j - 1).Value
          ElseIf j = 4 Or j = 5 Then
            xlSheet.Cells(Nextrow, j).Value = Format(TempExcelrs.Fields(j - 1).Value, "dd MMMM yyyy")
          ElseIf j = 6 Or j = 7 Or j = 8 Then
            xlSheet.Cells(Nextrow, j).Value = FormatCurrency(TempExcelrs.Fields(j - 1).Value, 2, True, True, True)
          ElseIf j = 9 Then
            xlSheet.Cells(Nextrow, j).Value = FormatCurrency(vProfit, 2, True, True, True)
          Else
            xlSheet.Cells(Nextrow, j).Value = TempExcelrs.Fields(j - 1).Value
          End If
          xlSheet.Cells(Nextrow, j).HorizontalAlignment = -4108
        Next j
        TempExcelrs.MoveNext
        Nextrow = Nextrow + 1
        vGotTotal = False
      Next i
      
      'print totals
      Nextrow = Nextrow + 1
      xlSheet.Cells(Nextrow, 6).Value = FormatCurrency(vTotals(2), 2, True, True, True)
      xlSheet.Cells(Nextrow, 6).HorizontalAlignment = -4108
      xlSheet.Cells(Nextrow, 7).Value = FormatCurrency(vTotals(3), 2, True, True, True)
      xlSheet.Cells(Nextrow, 7).HorizontalAlignment = -4108
      xlSheet.Cells(Nextrow, 8).Value = FormatCurrency(vTotals(1), 2, True, True, True)
      xlSheet.Cells(Nextrow, 8).HorizontalAlignment = -4108
      xlSheet.Cells(Nextrow, 9).Value = FormatCurrency(vTotals(4), 2, True, True, True)
      xlSheet.Cells(Nextrow, 9).HorizontalAlignment = -4108

      
    Case 11     'export postal codes
      TempExcelrs.MoveFirst
      For i = 1 To recCunt
        For j = 1 To fldCunt
          If j = 1 Then
            xlSheet.Cells(Nextrow, j).Value = "'" & TempExcelrs.Fields(j - 1).Value
          Else
            If j = 3 Then
              GetFile = Right(TempExcelrs.Fields(j - 1).Value, Len(TempExcelrs.Fields(j - 1).Value) - InStrRev(TempExcelrs.Fields(j - 1).Value, "\"))
              xlSheet.Cells(Nextrow, j).Value = GetFile
            Else
              xlSheet.Cells(Nextrow, j).Value = TempExcelrs.Fields(j - 1).Value
              End If
          End If
          xlSheet.Cells(Nextrow, j).HorizontalAlignment = -4108
        Next j
        TempExcelrs.MoveNext
        Nextrow = Nextrow + 1
      Next i
      
      
  End Select
  Screen.MousePointer = vbNormal
  
End With

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
  Case 1       ' single veh details, called from addvehicle
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
   Case 2       ' deal List
     ReportHead = "Vehicle Details"
     xlSheet.Range("A1").ColumnWidth = 6
     xlSheet.Range("B1").ColumnWidth = 16
     xlSheet.Range("C1").ColumnWidth = 16
     xlSheet.Range("D1").ColumnWidth = 14
     xlSheet.Range("E1").ColumnWidth = 16
     xlSheet.Range("F1").ColumnWidth = 16
     xlSheet.Range("G1").ColumnWidth = 14
     xlSheet.Range("H1").ColumnWidth = 26
     xlSheet.Range("I1").ColumnWidth = 13
     xlSheet.Range("J1").ColumnWidth = 11
     xlSheet.Range("K1").ColumnWidth = 11
     xlSheet.Range("L1").ColumnWidth = 11
     xlSheet.Range("M1").ColumnWidth = 12
    Case 3    'client listing
     ReportHead = "Client Listing"
     xlSheet.Range("A1").ColumnWidth = 18
     xlSheet.Range("B1").ColumnWidth = 18
     xlSheet.Range("C1").ColumnWidth = 16
     xlSheet.Range("D1").ColumnWidth = 20
     xlSheet.Range("E1").ColumnWidth = 20
     xlSheet.Range("F1").ColumnWidth = 16
     xlSheet.Range("G1").ColumnWidth = 16
     xlSheet.Range("H1").ColumnWidth = 20
     xlSheet.Range("I1").ColumnWidth = 10
     xlSheet.Range("J1").ColumnWidth = 16
    Case 4    'Deal listing
     ReportHead = "Deal Listing"
     xlSheet.Range("A1").ColumnWidth = 6
     xlSheet.Range("B1").ColumnWidth = 10
     xlSheet.Range("C1").ColumnWidth = 16
     xlSheet.Range("D1").ColumnWidth = 18
     xlSheet.Range("E1").ColumnWidth = 18
     xlSheet.Range("F1").ColumnWidth = 16
     xlSheet.Range("G1").ColumnWidth = 16
     xlSheet.Range("H1").ColumnWidth = 16
     xlSheet.Range("I1").ColumnWidth = 16
    
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
  
