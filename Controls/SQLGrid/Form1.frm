VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SQL Grid"
   ClientHeight    =   4290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Caption         =   "search"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin Project1.SQLGrid SQLGrid1 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6165
   End
   Begin VB.Label lblSelect 
      Alignment       =   1  'Right Justify
      Caption         =   "No Item Selected"
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conn2 As ADODB.Connection

Private Sub cmdSearch_Click()

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Text1.SetFocus
If SQLGrid1.LoadGrid(Text1.Text) = False Then MsgBox SQLGrid1.Error: Exit Sub

End Sub

Private Sub Form_Load()

  Set Conn2 = New ADODB.Connection
  Conn2.ConnectionString = "Provider=microsoft.jet.OLEDB.4.0; ;" & "Data Source = " & "[ YOUR CONNECTION DETAILS ]"
  Conn2.Open
  If SQLGrid1.StartGrid("Select * From PriShipment", Conn2, "Customer") = False Then MsgBox SQLGrid1.Error: Exit Sub
  If SQLGrid1.AddSearchField("Customer") = False Then MsgBox SQLGrid1.Error: Exit Sub
  If SQLGrid1.AddSearchField("ID") = False Then MsgBox SQLGrid1.Error: Exit Sub
  If SQLGrid1.AddSearchField("SapAcc") = False Then MsgBox SQLGrid1.Error: Exit Sub
  If SQLGrid1.AddSearchField("Invoice") = False Then MsgBox SQLGrid1.Error: Exit Sub
  SQLGrid1.FieldHeading(2) = "Customer Name"
  If SQLGrid1.LoadGrid = False Then MsgBox SQLGrid1.Error: Exit Sub

End Sub

Private Sub Form_Resize()

  SQLGrid1.Move 120, 600, Me.ScaleWidth - 240, Me.ScaleHeight - 720

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set Conn2 = Nothing

End Sub

Private Sub SQLGrid1_Click(ID As Long)

lblSelect.Caption = "Selected ID: " & ID

End Sub

Private Sub SQLGrid1_DblClick(ID As Long)

MsgBox ID

End Sub

Private Sub SQLGrid1_NoSelection()

lblSelect.Caption = "No Item Selected"

End Sub
