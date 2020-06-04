VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Disconnect"
      Height          =   495
      Left            =   9480
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load"
      Height          =   495
      Left            =   9480
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   495
      Left            =   8280
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents clsSql As clsDB
Attribute clsSql.VB_VarHelpID = -1

Private Sub clsSql_Connected(Status As Boolean)
  
  Label2.Caption = "Connected : " & Status & vbCrLf & _
                    "Datebase Name : " & clsSql.DatabaseName & vbCrLf & _
                    "Datebase Path : " & clsSql.DatabasePath
                    
End Sub

Private Sub Command1_Click()

  If clsSql.OpenConnection("DB.MDB", App.Path, "") = False Then
    MsgBox clsSql.Error
  End If
  
End Sub

Private Sub Command2_Click()

  Dim iTmx As Object

  ListView1.ListItems.Clear
  ListView1.ColumnHeaders.Clear
  ListView1.ColumnHeaders.Add 1, , "ID"
  ListView1.ColumnHeaders.Add 2, , "Cat Type"
  ListView1.ColumnHeaders.Add 3, , "Code"
  ListView1.ColumnHeaders.Add 4, , "Description"
  ListView1.ColumnHeaders.Add 5, , "Blocked"
  ListView1.ColumnHeaders.Add 6, , "Deleted"
  
  If clsSql.OpenRecordSet("Select * from Category") = False Then
    MsgBox clsSql.Error
    Exit Sub
  End If
  
  Do While Not clsSql.Recordset.EOF
    Set iTmx = ListView1.ListItems.Add(, , clsSql.Recordset.Fields("ID").Value)
    iTmx.SubItems(1) = clsSql.Recordset.Fields(1).Value
    iTmx.SubItems(2) = clsSql.Recordset.Fields(2).Value
    iTmx.SubItems(3) = clsSql.Recordset.Fields(3).Value
    iTmx.SubItems(4) = clsSql.Recordset.Fields(4).Value
    iTmx.SubItems(5) = clsSql.Recordset.Fields(5).Value
    clsSql.Recordset.MoveNext
  Loop
  
  Label1.Caption = "Record Count : " & clsSql.RecordCount
  clsSql.CloseRecordset

End Sub

Private Sub Command3_Click()

  clsSql.CloseConnection
  
End Sub

Private Sub Form_Load()

  Set clsSql = New clsDB
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set clsSql = Nothing
  
End Sub

