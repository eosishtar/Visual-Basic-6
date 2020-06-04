VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl SQLGrid 
   ClientHeight    =   5520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7980
   ScaleHeight     =   5520
   ScaleWidth      =   7980
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   6960
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid GRD 
      Height          =   4335
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7646
      _Version        =   393216
      SelectionMode   =   1
   End
End
Attribute VB_Name = "SQLGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim IsLastItem2 As Boolean
Dim IsFirstItem2 As Boolean
Dim LastAsc As Boolean
Dim PrevSortCo As Integer
Dim LastIDClicked As Long
Dim ButtonClicked As Long
Dim HeaderClicked As Boolean
Dim SearchBit As String
Dim SearchFields As String
Dim WhereString As String
Dim GridOrder As String
Dim GridBoolean As Boolean
Dim GridTmpStr As String
Dim GridCounter As Integer
Private Type GridFieldType
  FieldName As String
  IsSearchField As Boolean
  Heading As String
End Type
Private Type GridInfoType
  GridFields As Integer
  GridField(1 To 256) As GridFieldType
  GridBaseSQL As String
  GridConn As ADODB.Connection
  GridOrderBy As String
  GridSQL As ADODB.Recordset
  GridInfoPopulated As Boolean
End Type
Dim gError As String
Dim GI As GridInfoType
Public Event Click(ID As Long)
Public Event DblClick(ID As Long)
Public Event NoSelection()

Private Sub SortFlex(TheCol As Integer, ForceDesc As Boolean, ParamArray IsString() As Variant)

  Dim i
  Dim Headline
  Dim Ascend
  Dim Decend
  
  Screen.MousePointer = 11
  On Error Resume Next
  GRD.Col = TheCol
  For i = 0 To GRD.Cols - 1
    Headline = GRD.TextMatrix(0, i)
    Ascend = Right$(Headline, 1) = "+"
    Decend = Right$(Headline, 1) = "-"
    If Ascend Or Decend Then Headline = Left$(Headline, Len(Headline) - 1)
    If i = TheCol Then
      If Ascend Or ForceDesc Then
        LastAsc = True
        GRD.TextMatrix(0, i) = Headline & "-"
        If IsMissing(IsString(i)) Then
        GRD.Sort = flexSortGenericDescending
      Else
        If IsString(i) Then
          GRD.Sort = flexSortStringDescending
        Else
          GRD.Sort = flexSortNumericDescending
        End If
      End If
    Else
      LastAsc = False
      GRD.TextMatrix(0, i) = Headline & "+"
      If IsMissing(IsString(i)) Then
        GRD.Sort = flexSortGenericAscending
      Else
        If IsString(i) Then
          GRD.Sort = flexSortStringAscending
        Else
          GRD.Sort = flexSortNumericAscending
        End If
      End If
    End If
    Else
      GRD.TextMatrix(0, i) = Headline
    End If
  Next i
  On Error GoTo 0
  Screen.MousePointer = 0

End Sub

Private Sub HighlightGridRow(iRow As Long)

  With GRD
    If .Rows > 1 Then
      .Row = iRow
      .Col = 1
      .ColSel = .Cols - 1
      .RowSel = iRow
    End If
  End With

End Sub

Private Sub SetGridColumnWidth()
    
  Dim InnerLoopCount As Long
  Dim OuterLoopCount As Long
  Dim lngLongestLen As Long
  Dim sLongestString As String
  Dim lngColWidth As Long
  Dim szCellText As String

  For OuterLoopCount = 0 To GRD.Cols - 1
    sLongestString = ""
    lngLongestLen = 0
    For InnerLoopCount = 0 To GRD.Rows - 1
      szCellText = GRD.TextMatrix(InnerLoopCount, OuterLoopCount)
      If Len(szCellText) > lngLongestLen Then
        lngLongestLen = Len(szCellText)
        sLongestString = szCellText
      End If
    Next
    lngColWidth = Picture1.TextWidth(sLongestString)
    GRD.ColWidth(OuterLoopCount) = lngColWidth + 200
  Next
  
End Sub

Private Sub SetGrid()

GRD.Move 0, 0, UserControl.Width, UserControl.Height
If GRD.Width <> UserControl.Width Then UserControl.Width = GRD.Width
If GRD.Height <> UserControl.Height Then UserControl.Height = GRD.Height

End Sub

Public Property Get FieldHeading(FieldNum As Integer) As String

FieldHeading = GI.GridField(FieldNum).Heading

End Property

Public Property Let FieldHeading(FieldNum As Integer, ByVal vNewValue As String)

GI.GridField(FieldNum).Heading = vNewValue

End Property

Private Sub GRD_Click()

If ButtonClicked <> 1 Then Exit Sub
If HeaderClicked = True Then
    PrevSortCo = GRD.MouseCol
    SortFlex GRD, GRD.MouseCol, False, False, True, True, True
    IsUnselected
Else
    If GRD.MouseRow > 0 Then
        If GRD.Rows = 1 Then
            IsUnselected
        Else
            IsSelected
        End If
    Else
        IsUnselected
    End If
End If

End Sub

Private Sub IsUnselected()

If LastIDClicked = 0 Then Exit Sub
IsFirstItem2 = False
IsLastItem2 = False
LastIDClicked = 0
RaiseEvent NoSelection

End Sub

Private Sub IsSelected()

Dim VV As Long
GRD.Col = 0
VV = Val(GRD.Text)
If LastIDClicked = VV Then Exit Sub
LastIDClicked = VV
If VV > 0 Then RaiseEvent Click(VV)
If GRD.Row = 1 Then IsFirstItem2 = True
If GRD.Row = GRD.Rows - 1 Then IsLastItem2 = True
HighlightGridRow GRD.Row

End Sub

Private Sub GRD_DblClick()

Dim VV2 As Integer
If HeaderClicked = True Then Exit Sub
VV2 = GRD.MouseCol + 1
GRD.Col = 0
If Val(GRD.Text) > 0 Then RaiseEvent DblClick(Val(GRD.Text))

End Sub

Private Sub GRD_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

ButtonClicked = Button
If y < GRD.RowHeight(0) Then HeaderClicked = True Else HeaderClicked = False

End Sub

Private Sub UserControl_Initialize()

GI.GridInfoPopulated = False
SetGrid

End Sub

Private Sub UserControl_Resize()

SetGrid

End Sub

Private Sub UserControl_Terminate()

Set GI.GridSQL = Nothing
Set GI.GridConn = Nothing

End Sub

Public Function Error() As String

  Error = gError

End Function

Private Function GetFieldsFromSQL() As Boolean

  GetFieldsFromSQL = False
  GI.GridFields = 0
  On Error Resume Next
  Set GI.GridSQL = New ADODB.Recordset
  If GI.GridSQL.State = 1 Then Set GI.GridSQL = Nothing
  GI.GridSQL.Open GI.GridBaseSQL & " where ID = 0", GI.GridConn, adOpenStatic, adLockReadOnly
  If Err <> 0 Then gError = "Could not GetFieldsFromSQL: " & Err.Description: On Error GoTo 0: Set GI.GridSQL = Nothing: Exit Function
  For GridCounter = 0 To GI.GridSQL.Fields.Count - 1
    GI.GridFields = GI.GridFields + 1
    GI.GridField(GI.GridFields).FieldName = UCase(GI.GridSQL.Fields(GridCounter).Name)
    GI.GridField(GI.GridFields).IsSearchField = False
    GI.GridField(GI.GridFields).Heading = GI.GridField(GI.GridFields).FieldName
  Next GridCounter
  GI.GridSQL.Close
  If Err <> 0 Then gError = "Could not GetFieldsFromSQL: " & Err.Description: On Error GoTo 0: Set GI.GridSQL = Nothing: Exit Function
  On Error GoTo 0
  Set GI.GridSQL = Nothing
  GetFieldsFromSQL = True

End Function

Public Function AddSearchField(SearchFieldToAdd As String) As Boolean

  AddSearchField = False
  If GI.GridInfoPopulated = False Then
    gError = "Cannot add search field " & SearchFieldToAdd & ". Grid not Started"
    Exit Function
  End If
  GridBoolean = False
  For GridCounter = 1 To GI.GridFields
    If GI.GridField(GridCounter).FieldName = UCase(SearchFieldToAdd) Then
      GridBoolean = True
      Exit For
    End If
  Next GridCounter
  If GridBoolean = False Then
    gError = "Field " & SearchFieldToAdd & " not found"
  Else
    GI.GridField(GridCounter).IsSearchField = True
    AddSearchField = True
  End If

End Function

Public Function StartGrid(SQLBaseString As String, MainConn As ADODB.Connection, OrderByString As String) As Boolean

  StartGrid = False
  GI.GridInfoPopulated = False
  GI.GridBaseSQL = UCase(SQLBaseString)
  If InStr(1, GI.GridBaseSQL, "WHERE") > 0 Then GI.GridBaseSQL = Left(GI.GridBaseSQL, InStr(1, GI.GridBaseSQL, "WHERE") - 1)
  GI.GridBaseSQL = Trim(GI.GridBaseSQL)
  Set GI.GridConn = MainConn
  If GetFieldsFromSQL = False Then Set GI.GridConn = Nothing: Exit Function
  GI.GridOrderBy = UCase(Trim(OrderByString))
  GI.GridInfoPopulated = True
  StartGrid = True

End Function

Private Function GetWhere(Optional SearchString As String) As String

If GI.GridOrderBy = "" Then
  GridOrder = ""
Else
  If InStr(1, GI.GridOrderBy, "ORDER BY") > 0 Then
    GridOrder = " " & GI.GridOrderBy
  Else
    GridOrder = " ORDER BY " & GI.GridOrderBy
  End If
End If
If SearchString = "" Then
  GetWhere = GI.GridBaseSQL & GridOrder
Else
  SearchFields = ""
  For GridCounter = 1 To GI.GridFields
    SearchBit = "Instr(1,Ucase(" & GI.GridField(GridCounter).FieldName & ")," & Chr(34) & SearchString & Chr(34) & ") > 0"
    If GI.GridField(GridCounter).IsSearchField = True Then
      If SearchFields = "" Then
        SearchFields = SearchBit
      Else
        SearchFields = SearchFields & " OR " & SearchBit
      End If
    End If
  Next GridCounter
  If SearchFields = "" Then
    GetWhere = GI.GridBaseSQL & GridOrder
  Else
    GetWhere = GI.GridBaseSQL & " WHERE " & SearchFields & " " & GridOrder
  End If
End If

End Function

Private Sub LoadRecordsetIntoGrid(rs As Recordset, Optional AutosizeColumns As Boolean = True, Optional HighlightFirstRow As Boolean = True)
                
  Dim x As Long
  Dim Count As Long

  GRD.Redraw = False
  GRD.Clear
  GRD.Rows = 2
  GRD.FixedRows = 1
  GRD.Row = 0
  GRD.Cols = rs.Fields.Count + 1
  For x = 0 To rs.Fields.Count - 1
    GRD.Col = x + 1
    GRD.Text = rs.Fields(x).Name
    GRD.ColData(x + 1) = rs.Fields(x).Type
  Next

  If rs.CursorLocation = adUseClient Then
    GRD.Rows = rs.RecordCount + 1
    For Count = 1 To rs.RecordCount
      For x = 0 To rs.Fields.Count - 1
        GRD.TextMatrix(Count, x) = "" & CVar(rs.Fields(x).Value)
      Next
      rs.MoveNext
    Next
  ElseIf rs.CursorLocation = adUseServer Then
    Do While Not rs.EOF
      Count = Count + 1
      If Count >= GRD.Rows Then GRD.Rows = GRD.Rows + 100
      For x = 0 To rs.Fields.Count - 1
        GRD.TextMatrix(Count, x) = "" & CVar(rs.Fields(x).Value)
      Next
      rs.MoveNext
    Loop
    GRD.Rows = Count + 1
  End If
  If AutosizeColumns Then SetGridColumnWidth
  If HighlightFirstRow Then
      If GRD.Rows > 1 Then HighlightGridRow 1
  End If
  GRD.Redraw = True

End Sub

Public Function LoadGrid(Optional SearchString As String, Optional AutosizeColumns As Boolean = True, Optional HighlightFirstRow As Boolean = True, Optional ShowIDColumn As Boolean = False) As Boolean

  LoadGrid = False
  If GI.GridInfoPopulated = False Then
    gError = "Cannot Load Grid. Grid not Started"
    Exit Function
  End If
  WhereString = GetWhere(UCase(SearchString))
  On Error Resume Next
  Set GI.GridSQL = New ADODB.Recordset
  If GI.GridSQL.State = 1 Then Set GI.GridSQL = Nothing
  GI.GridSQL.Open WhereString, GI.GridConn, adOpenStatic, adLockReadOnly
  If Err <> 0 Then gError = "Could not LoadGrid: " & Err.Description: On Error GoTo 0: Set GI.GridSQL = Nothing: Exit Function
  On Error GoTo 0
  LoadRecordsetIntoGrid GI.GridSQL, AutosizeColumns, HighlightFirstRow
  If ShowIDColumn = False Then GRD.ColWidth(0) = 0
  LoadGrid = True
  
End Function
