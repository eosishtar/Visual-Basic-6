VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Report "
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   13830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   9840
      TabIndex        =   35
      Top             =   7920
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Index           =   3
      Left            =   120
      TabIndex        =   32
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdDelete 
      Cancel          =   -1  'True
      Caption         =   "&Delete"
      Height          =   525
      Left            =   1320
      Picture         =   "frmCustReports.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Delete the Record"
      Top             =   7800
      Width           =   1125
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   525
      Left            =   120
      Picture         =   "frmCustReports.frx":2CE4
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Close"
      Top             =   7800
      Width           =   1125
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   525
      Left            =   12480
      Picture         =   "frmCustReports.frx":59BF
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Save Record"
      Top             =   7800
      Width           =   1245
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Query"
      Enabled         =   0   'False
      Height          =   525
      Left            =   11160
      Picture         =   "frmCustReports.frx":8632
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "View the Results"
      Top             =   7800
      Width           =   1245
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   525
      Left            =   11160
      Picture         =   "frmCustReports.frx":B4D9
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Back"
      Top             =   7800
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Caption         =   "Choose Fields "
      Height          =   6255
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   11655
      Begin VB.CommandButton cmdSwapFrame 
         Height          =   400
         Left            =   5640
         Picture         =   "frmCustReports.frx":E282
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Filter Options"
         Top             =   5160
         Width           =   400
      End
      Begin VB.Frame fraSave 
         Height          =   855
         Left            =   6360
         TabIndex        =   19
         Top             =   120
         Width           =   5175
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   21
            Text            =   "My Report"
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label9 
            Caption         =   "Report Name"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   375
            Width           =   975
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Width           =   4335
      End
      Begin VB.ListBox List1 
         Height          =   3885
         Index           =   0
         Left            =   480
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   1680
         Width           =   4935
      End
      Begin VB.ListBox List1 
         Height          =   3210
         Index           =   1
         Left            =   6360
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   1680
         Width           =   4935
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   400
         Left            =   5640
         Picture         =   "frmCustReports.frx":11141
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Add Item"
         Top             =   2520
         Width           =   400
      End
      Begin VB.CommandButton cmdRemove 
         Height          =   400
         Left            =   5640
         Picture         =   "frmCustReports.frx":13FE8
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Remove Item"
         Top             =   3240
         Width           =   400
      End
      Begin VB.CommandButton cmdReset 
         Height          =   400
         Left            =   5640
         Picture         =   "frmCustReports.frx":16D91
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Remove All"
         Top             =   4680
         Width           =   400
      End
      Begin VB.Frame Frame3 
         Caption         =   " Filter Options "
         Height          =   1095
         Left            =   6360
         TabIndex        =   23
         Top             =   5040
         Width           =   5175
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   650
            Width           =   3375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Enable Date Filter"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   24
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label11 
            Caption         =   "Order By"
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   675
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "  Info "
         Height          =   1095
         Left            =   6360
         TabIndex        =   2
         Top             =   5040
         Width           =   5175
         Begin VB.Label Label4 
            Caption         =   "Data Type"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   680
            Width           =   975
         End
         Begin VB.Label Label6 
            Height          =   255
            Left            =   1440
            TabIndex        =   5
            Top             =   680
            Width           =   3135
         End
         Begin VB.Label Label7 
            Height          =   255
            Left            =   1440
            TabIndex        =   4
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label8 
            Caption         =   "Table"
            Height          =   255
            Left            =   360
            TabIndex        =   3
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Table Name"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Fields in Table"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Selected Fields"
         Height          =   255
         Left            =   6120
         TabIndex        =   13
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Total Fields"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   5640
         Width           =   2895
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "  View Custom Report Results "
      Height          =   6255
      Left            =   2040
      TabIndex        =   27
      Top             =   1080
      Width           =   11655
      Begin MSComctlLib.ListView ListView1 
         Height          =   5295
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9340
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblRecords 
         Caption         =   "ListCount"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   5880
         Width           =   3735
      End
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Excel Report Builder"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   120
      TabIndex        =   34
      Top             =   240
      Width           =   13635
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reports"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   240
      TabIndex        =   33
      Top             =   1080
      Width           =   1515
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   14280
      Y1              =   7605
      Y2              =   7605
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   14400
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Index           =   0
      Begin VB.Menu mnuSubReport 
         Caption         =   "Report Builder"
         Index           =   0
      End
      Begin VB.Menu mnuSubReport 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuSubReport 
         Caption         =   "Add New"
         Index           =   2
      End
      Begin VB.Menu mnuSubReport 
         Caption         =   "-"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmCustReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cnt As Integer
Dim DIRTY_ As Boolean
Const LongString = "                                                                                                                    "


Private Sub cmdAdd_Click()
  
  DIRTY_ = True
  If Not CheckList(0) Then
    Exit Sub
  End If
  
  'add item to selected list
  For Cnt = 0 To List1(0).ListCount - 1
    If List1(0).Selected(Cnt) = True Then
      List1(1).AddItem Combo1.Text & "." & List1(0).List(Cnt)
    End If
    
    cmdSave.Enabled = True
    cmdTest.Enabled = True
  Next Cnt
  
  'remove from current list
  For Cnt = List1(0).ListCount - 1 To 0 Step -1
    If List1(0).Selected(Cnt) = True Then
      List1(0).RemoveItem Cnt
    End If
  Next Cnt
  
  Call DoFilterOptions
  Label5.Caption = "Total fields : " & List1(0).ListCount

End Sub

Private Sub cmdBack_Click()
  Frame2.Visible = True
  Frame5.Visible = False
  cmdTest.Visible = True
  cmdBack.Visible = False
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdDelete_Click()

Dim rEs As Integer

rEs = MsgBox("Are you sure you want to delete " & Rep.RepName & " ?", vbInformation + vbYesNo, "Confirm delete?")
  '...delete the custom report
  If rEs = 6 Then
  Set rsFields = New ADODB.Recordset
    sql = "Select * FROM [CustReports] WHERE ID = " & Rep.ID
      With rsFields
        .Open sql, cnDB, adOpenKeyset, adLockOptimistic
          If .EOF Then
            .Close
            MsgBox "An has error occurred.", vbCritical + vbOKOnly, "Delete failed."
          End If
          
          .Delete
          .Update
          .Close
                    
          MsgBox Rep.RepName & " was successfully deleted.", vbCritical + vbOKOnly, "Deleted"
      End With
  End If
  
  Call LoadReportList(List1(3))
  Set rsFields = Nothing

End Sub

Private Sub cmdRemove_Click()

  Dim vTable As String
  Dim vTable2 As String
  Dim vItem As String
  Dim vT As Integer
  
  DIRTY_ = True
  
  'add item to Current list
  For Cnt = 0 To List1(1).ListCount - 1
    If List1(1).Selected(Cnt) = True Then
      vTable = List1(1).List(Cnt)
      vTable2 = Left(vTable, Len(vTable) - 1 - Len(Right(vTable, Len(vTable) - InStrRev(vTable, "."))))
      If Trim$(vTable2) = Trim$(Combo1.Text) Then                                  'cant add item if not on the correct table
        vItem = Trim$(Left(Right(vTable, Len(vTable) - Len(vTable2) - 1), 30))     'remove table name, only add field with data type
        If AddToList(vItem) Then List1(0).AddItem vItem                            'must check that that the item doesnt already exisit
      End If
    End If
  Next Cnt
  
  'remove from Selected list
  For Cnt = List1(1).ListCount - 1 To 0 Step -1
    If List1(1).Selected(Cnt) = True Then
      List1(1).RemoveItem Cnt
    End If
  Next Cnt
  
  'disable the save option if nothing to save
  If List1(1).ListCount = 0 Then
    cmdSave.Enabled = False
    cmdTest.Enabled = False
  End If
  
  Call DoFilterOptions
  Label5.Caption = "Total fields : " & List1(0).ListCount
  
End Sub

Private Function AddToList(vItem As String) As Boolean
'this sub will check and make sure that you dont add duplicates to an existing listbox
  Dim dX As Integer
  
  AddToList = True
  For dX = 0 To List1(0).ListCount - 1
    If Trim$(vItem) = Trim$(Left(List1(0).List(dX), 30)) Then
      AddToList = False
      Exit For
    End If
  Next dX

End Function

Private Sub cmdReset_Click()

  DIRTY_ = True
  Combo1_Click
  List1(1).Clear
  Combo2.Clear
  
End Sub


Private Sub cmdSave_Click()
  Dim SavedSql  As String
  Dim vMsg As String
  
  vMsg = ""
  If Trim$(Text1.Text) = "" Then
    vMsg = " * Report name." & vbCrLf
  End If
  If List1(1).ListCount < 1 Then
    vMsg = vMsg & " * Select at least one field."
  End If
  If vMsg <> "" Then
    MsgBox "Please complete the following first " & vbCrLf & vMsg, vbInformation + vbOKOnly, Me.Caption
    Exit Sub
  End If
    

  
  ', vbInformation + vbOKOnly, Me.Caption
  
  SavedSql = GetFieldsFromList
  DoEvents

  Set rsFields = New ADODB.Recordset
  rsFields.Open "Select * from CustReports WHERE ID = " & Rep.ID, cnDB, adOpenKeyset, adLockOptimistic
    If rsFields.EOF Then
      rsFields.AddNew
    End If
      rsFields!Desc = Trim$(Text1.Text)
      rsFields!Record = SavedSql
    
    rsFields.Update
    rsFields.Close
    
  Set rsFields = Nothing
  
  List1(0).Clear
  List1(1).Clear
  Combo1.ListIndex = -1
  Label6.Caption = ""
  Label7.Caption = ""
  Text1.Text = ""
  DIRTY_ = False
  Call LoadReportList(List1(3))
  Call LoadReportMainMenu(True)
  
End Sub

Private Sub cmdSwapFrame_Click()

  If Frame1.Visible = True Then
    Call DoFilterOptions
    Frame1.Visible = False
    Frame3.Visible = True
    cmdSwapFrame.ToolTipText = "Info Options"
  Else
    Frame1.Visible = True
    Frame3.Visible = False
    cmdSwapFrame.ToolTipText = "Filter Options"
  End If

End Sub

Private Sub cmdTest_Click()
  Dim Temp_SQL As String
  Dim Temp2_SQL As String
  
  Frame5.Visible = True
  Frame2.Visible = False
  cmdTest.Visible = False
  cmdBack.Visible = True
  
  Temp_SQL = GetFieldsFromList          'Get the string
  If List1(1).ListCount > 1 Then
    Temp2_SQL = MakeCustomSQL(Temp_SQL)   'show the results
    LoadReport (Temp2_SQL)
  Else
    MsgBox "Please select as least two field before running the report.", vbInformation + vbOKOnly, Me.Caption
    Frame5.Visible = False
    Frame2.Visible = True
    cmdTest.Visible = True
    cmdBack.Visible = False
  End If
  
End Sub


Private Sub Form_Load()


' -------------------------       TEMP STUFF   ---------------------------------------
vPassword = "Starlight1"
daDb = "C:\Users\Mark.Lang\Desktop\Snippets\_WORK\CustomReport\DB.mdb"
cmd = "Provider=microsoft.jet.OLEDB.4.0; ;" & "Data Source = " & daDb & ";Jet OLEDB:Database Password=" & vPassword

'...Open Connection
With cnDB
  .Provider = "Microsoft.Jet.OLEDB.4.0"
  .Open cmd
End With

'must set this in form load
EXCLUDE_TABLES = "CustReports"
EXCLUDE_FIELDS = ""

' -------------------------       TEMP STUFF   ---------------------------------------



  Call LoadReportList(List1(3))    'load menu
  Call LoadReportMainMenu          'load main menu reprot
  Call ListTables                  'List tables
  Frame3.Visible = False           'Filter Options
  Frame5.Visible = False           'View Results
  
End Sub

Private Sub Command1_Click()
  
  frmCustRepSelector.Show

End Sub

Private Sub Combo1_Click()

  DIRTY_ = True
  If Combo1.ListIndex >= 0 Then
    ListFields (Combo1.Text)
    Label2.Caption = "Showing fields for table '" & Combo1.Text & "'"
  End If

End Sub

Private Sub ListTables()
  
  '...open the schema
  Set rsTables = cnDB.OpenSchema(adSchemaTables)

  With rsTables
    If Not .EOF Then
        .Filter = "TABLE_TYPE = 'TABLE'"
        Do While Not .EOF
          If InStr(1, EXCLUDE_TABLES, .Fields("TABLE_NAME")) <= 0 Then
            Combo1.AddItem .Fields("TABLE_NAME")
          End If
          .MoveNext
        Loop
    End If
  End With
  
  rsTables.Close
  Set rsTables = Nothing
    
End Sub

 Private Sub ListFields(ByVal db_table_name As String)

  List1(0).Clear
   
  '...get the table names.
  Set rsFields = New ADODB.Recordset
  
  rsFields.Open "select * from " & db_table_name, cnDB, adOpenKeyset, adLockReadOnly
  Cnt = 1
  
  For Each tField In rsFields.Fields
    If InStr(1, EXCLUDE_FIELDS, tField.Name) <= 0 Then
      'tType = GetDataType(tField.Type)
      List1(0).AddItem tField.Name & LongString & tField.Type
      Cnt = Cnt + 1
    End If
  Next tField
  
  Label5.Caption = "Total fields : " & Cnt
  rsFields.Close
  Set rsFields = Nothing
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim mRsp As VbMsgBoxResult

  If DIRTY_ Then
    mRsp = MsgBox("Exit without saving?", vbQuestion + vbYesNo, Me.Caption)
    If mRsp <> vbYes Then
      Cancel = 1
      Exit Sub
    End If
  End If
  

cnDB.Close
Set cnDB = Nothing

End Sub

Private Sub List1_Click(Index As Integer)
  Dim VS As String
  Dim Temp_Item As String
  
  Label6.Caption = ""
  Label7.Caption = ""
  
  Select Case Index
    Case 0     'listed fields
         
    Case 1     'selected fields
      'display captions
      Label6.Caption = Trim$(Right(List1(1).List(List1(1).ListIndex), 30))
      Label6.Caption = GetDataType(Val(Label6.Caption))
      
      VS = List1(1).List(List1(1).ListIndex)
      Label7.Caption = Left(VS, Len(VS) - 1 - Len(Right(VS, Len(VS) - InStrRev(VS, "."))))
      
    Case 3    'menu selector
      Frame2.Visible = True
      Frame5.Visible = False
      cmdBack.Visible = False
      cmdTest.Visible = True
    
      Select Case List1(3).ListIndex
        Case 0        'new
          List1(0).Clear
          List1(1).Clear
          Combo1.ListIndex = -1
          Label6.Caption = ""
          Label7.Caption = ""
          Text1.Text = ""
          cmdSave.Caption = "Add"
          cmdDelete.Enabled = False
          Rep.ID = 0
        
        Case 1, 2    'nothing
        
        Case Else
          Rep.ID = Val(Right(List1(3).List(List1(3).ListIndex), 3))
          cmdSave.Caption = "Save"
          cmdDelete.Enabled = True
          If Rep.ID <> 0 Then
            GetReportSetting Rep.ID     'fetch the report detail
            Text1.Text = Rep.RepName
            If Rep.DateFilter Then Check1(0).Value = 1 Else Check1(0).Value = 0
            Combo2.ListIndex = -1       'SetComboText    'STILL TO DO
            
            List1(1).Clear
            For Cnt = 1 To Rep.FieldCount
            Temp_Item = Rep.DataFields(Cnt).TableName & "." & Rep.DataFields(Cnt).FieldName & LongString & Rep.DataFields(Cnt).DateType
            List1(1).AddItem Temp_Item
            cmdSave.Enabled = True
            cmdTest.Enabled = True
            Next Cnt
          End If
          DoFilterOptions
      End Select
      
      DIRTY_ = False
  End Select

End Sub

Private Sub List1_DblClick(Index As Integer)
Dim vTable As String
Dim vTable2 As String
Dim vItem As String

  Select Case Index
  
    Case 0
      If Not CheckList(0) Then
        Exit Sub
      End If
      
      'add item to selected list
      List1(1).AddItem Combo1.Text & "." & List1(0).List(List1(0).ListIndex)
      cmdSave.Enabled = True
      cmdTest.Enabled = True
      'remove from current list
      List1(0).RemoveItem List1(0).ListIndex
    
    Case 1
      'add item to Current list
      vTable = List1(1).List(List1(1).ListIndex)
      vTable2 = Left(vTable, Len(vTable) - 1 - Len(Right(vTable, Len(vTable) - InStrRev(vTable, "."))))
      If Trim$(vTable2) = Trim$(Combo1.Text) Then                                  'cant add item if not on the correct table
        vItem = Trim$(Left(Right(vTable, Len(vTable) - Len(vTable2) - 1), 30))     'remove table name, only add field with data type
        If AddToList(vItem) Then List1(0).AddItem vItem                            'must check that that the item doesnt already exisit
      End If
      
      'remove from Selected list
      List1(1).RemoveItem List1(1).ListIndex
      
  End Select
  
DIRTY_ = True


Call DoFilterOptions
Label5.Caption = "Total fields : " & List1(0).ListCount

End Sub

Private Function CheckList(vIndex As Integer) As Boolean

  CheckList = True
  If List1(vIndex).ListIndex = -1 Or List1(vIndex).ListCount < 1 Then
    CheckList = False
  End If

End Function

Private Function GetFieldsFromList() As String

  Dim tItem As String
  Dim tType As String
  Dim tSetting As String
  
  GetFieldsFromList = ""
  GetFieldsFromList = List1(1).ListCount & "," & Check1(0).Value & "," & Combo2.Text & "|"
  
  'Get all the fields from the selected list
  For Cnt = 0 To List1(1).ListCount
  tItem = Trim$(Left(List1(1).List(Cnt), 50))
  tType = Trim$(Right(List1(1).List(Cnt), 10))
    If Cnt = 0 Then
      GetFieldsFromList = GetFieldsFromList & tItem & "." & tType
    Else
      If tItem <> "" Then
        GetFieldsFromList = GetFieldsFromList & ", " & tItem & "." & tType
      End If
    End If
  Next Cnt
  
  '...setting string will be returned as follows
  '...Nr Selected Fields + DatePicker yes/no + Order by     AND SELECTED FIELDS  (TableName,FieldName, DataType)
  
End Function

Private Sub DoFilterOptions()
  Dim GG As Integer
  Dim gItem As String
  
  'first check if can enable date filter. must have selected at least one date field
  Check1(0).Enabled = False
  For GG = 1 To List1(1).ListCount
    gItem = UCase(Trim$(Left(List1(1).List(GG - 1), 30)))
    If InStr(1, gItem, "DATE") > 0 Then
      Check1(0).Enabled = True
      Exit For
    End If
  Next GG
  
  Combo2.Clear
  Combo2.AddItem ""
  For GG = 1 To List1(1).ListCount
    gItem = Trim$(Left(List1(1).List(GG - 1), 30))
    If gItem <> "" Then
      Combo2.AddItem gItem
    End If
  Next GG
  
  'disable the save option if nothing to save
  If List1(1).ListCount = 0 Then
    cmdSave.Enabled = False
    cmdTest.Enabled = False
  End If

End Sub

Private Function LoadReport(vCustomSQL As String)

  Dim sString As String
  Dim itmx As Object
  Dim i, j, k   As Integer
  Dim vTotRec As Integer

  lblRecords.Caption = ""
  DoEvents
  
  Set TempRs = New ADODB.Recordset
  
  With TempRs
    .Open vCustomSQL, cnDB, adOpenKeyset, adLockOptimistic

    If TempRs.EOF Then
        TempRs.Close
      Exit Function
    End If
    
    
    TempRs.MoveLast
    On Error GoTo ErrHandler:
    vTotRec = TempRs.RecordCount
    TempRs.MoveFirst
   
    '...set listview parameteTempRS
    ListView1.ColumnHeaders.Clear
    ListView1.ListItems.Clear
    ListView1.View = lvwReport
    ListView1.BorderStyle = ccFixedSingle
    ListView1.FullRowSelect = True
    ListView1.GridLines = True
    ListView1.LabelEdit = lvwManual

    '   count the columns and add them to the listview
    For i = 0 To TempRs.Fields.Count - 1
      ListView1.ColumnHeaders.Add , , TempRs.Fields(i).Name
    Next i
    
    '   count the rows and add the items and subitems
      TempRs.MoveFirst
      For j = 1 To vTotRec
        Set itmx = ListView1.ListItems.Add(, , TempRs.Fields(0).Value)
          For k = 1 To ListView1.ColumnHeaders.Count - 1
            On Error Resume Next
            If InStr(1, UCase(TempRs.Fields(k).Name), "DATE") > 0 Then
              itmx.SubItems(k) = Format(TempRs.Fields(k).Value, "dd MMMM yyyy")
            Else
              itmx.SubItems(k) = TempRs.Fields(k).Value
            End If
            On Error GoTo 0
          Next k
        TempRs.MoveNext
               
      Next j
      TempRs.Close
    End With
    
    'STILL TO DO
    'AltLVBackground ListView1, frmCustReports
    lblRecords.Caption = ListView1.ListItems.Count & " records loaded..."
    'Call AutosizeColumns(ListView1)     ' resize all the columns
     Set TempRs = Nothing
    
ErrHandler:
 Screen.MousePointer = vbNormal
 If Err.Number <> 0 Then
    MsgBox "An error has occurred and the results could not be displayed" & vbCrLf & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation + vbOKOnly, Me.Caption
    Err.Clear
    Set TempRs = Nothing
 End If

On Error GoTo 0
End Function


Private Sub mnuSubReport_Click(Index As Integer)
  Dim Cnt As Integer
  Dim Temp_Item As String
  
  Select Case Index
    Case 0
      MsgBox "Show Custom Report"
    
    Case 2
      MsgBox "Add New"
    
    Case Else
      Rep.ID = Val(mnuSubReport(Index).Tag)
      cmdSave.Caption = "Save"
      cmdDelete.Enabled = True
      
      If Rep.ID <> 0 Then
        GetReportSetting Rep.ID     'fetch the report detail
        frmCustRepSelector.Show
        '
        'Text1.Text = Rep.RepName
        'Check1(0).Value = Val(Rep.DateFilter)
        'Combo2.ListIndex = -1       'SetComboText    'STILL TO DO

        'List1(1).Clear
        'For Cnt = 1 To Rep.FieldCount
        'Temp_Item = Rep.DataFields(Cnt).TableName & "." & Rep.DataFields(Cnt).FieldName & LongString & Rep.DataFields(Cnt).DateType
        'List1(1).AddItem Temp_Item
        'cmdSave.Enabled = True
        'cmdTest.Enabled = True
        'Next Cnt
      End If
    
  End Select
  
  
  
End Sub






