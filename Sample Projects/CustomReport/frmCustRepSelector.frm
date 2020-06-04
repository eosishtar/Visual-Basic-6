VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCustRepSelector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter Options"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   5655
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   915
         Width           =   4455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   221446145
         CurrentDate     =   42555
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   221446145
         CurrentDate     =   42555
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   165
         Left            =   2760
         TabIndex        =   9
         Top             =   435
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   165
         Left            =   240
         TabIndex        =   8
         Top             =   435
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Filter "
         Height          =   165
         Left            =   240
         TabIndex        =   7
         Top             =   990
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Run"
      Height          =   525
      Left            =   4560
      Picture         =   "frmCustRepSelector.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Save Record"
      Top             =   3480
      Width           =   1245
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   525
      Left            =   120
      Picture         =   "frmCustRepSelector.frx":2C73
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Close"
      Top             =   3480
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   5655
      Begin VB.OptionButton Option3 
         Caption         =   "Not Selected"
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "False"
         Height          =   255
         Left            =   1980
         TabIndex        =   14
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "True"
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   -1200
      X2              =   13200
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   -1200
      X2              =   13080
      Y1              =   3285
      Y2              =   3285
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Options "
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
      Height          =   645
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   5835
   End
End
Attribute VB_Name = "frmCustRepSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const LongString = "                                                                                                                    "
Dim vDataType As String


Private Sub cmdClose_Click()
  Unload Me
End Sub


Private Sub cmdSave_Click()
  Dim vCnt As Integer
  Dim vSql As String
  Dim strSQL As String
  Dim vTables As String
  Dim vClnString As String
    
  vSql = ""
  vTables = ""
  'vClnString = cleanSQL(Text1.text)     'STILL TO DO
    
  'if user has selected date filter
  If Rep.DateFilter Then
    If Combo1.ListIndex <= 0 Then
      MsgBox "Please select the field to filter by date first.", vbInformation + vbOKOnly, Me.Caption
      Exit Sub
    End If
  End If

  For vCnt = 1 To Rep.FieldCount
    vSql = vSql & Rep.DataFields(vCnt).TableName & "." & Rep.DataFields(vCnt).FieldName & ", "
    vTables = vTables & Rep.DataFields(vCnt).TableName & ","     'collect all the table names
  Next vCnt
  vSql = Left(vSql, Len(vSql) - 2)
  
  'get the chosen columns
  vTables = GetSelectedTables(Rep.FieldCount, vTables)
  
  If Rep.DateFilter Then
    'has opted for date filter                                                  x = need to get date val from dtpicker
    'strSQL = "Select " & vSql & " FROM " & vTables & " WHERE " & Combo1.Text <= x And Combo1.Text >= x
  Else
    'no date filter set, but using custom filter
    If Combo1.ListIndex <= 0 Then
      strSQL = "Select " & vSql & " FROM " & vTables
    Else
      strSQL = "Select " & vSql & " FROM " & vTables & " WHERE " & Combo1.Text & " = " & vClnString
    End If
  End If
  
  


  Call ExcelReportDump(strSQL)

  
End Sub

Private Sub Combo1_Click()
  
  vDataType = ""
  Frame2.Visible = False      'text box
  Frame3.Visible = False      'radio buttons
  Text1.Text = ""
  Option1.Value = 0
  Option2.Value = 0
  Option3.Value = 0
  
    
  vDataType = Trim$(Right(Combo1.Text, 5))
  If vDataType <> "" Then
    Select Case vDataType
      Case 11     'Boolean
        Frame3.Visible = True
      Case 6, 3, 5 'Numbers Only
        Frame2.Visible = True
      Case 13     'Unknown
      Case 200, 201, 20, 202, 203 'String"
        If vDataType = "202" Then
          If InStr(1, UCase(Combo1.Text), "DATE") > 0 Then
            Frame2.Visible = False
          Else
            Frame2.Visible = True
          End If
        Else
          Frame2.Visible = True
        End If
      Case 205    'OLE Object
      Case Else
    End Select
  End If

End Sub

Private Sub Form_Load()

  DTPicker1.Value = Now()
  DTPicker2.Value = Now()
  Call LoadTableOptions
  
  If Rep.DateFilter Then
    DTPicker1.Enabled = True
    DTPicker2.Enabled = True
    'must set combo box to date fields
    
  Else
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
  End If

End Sub

Private Sub LoadTableOptions()
  Dim vCnt As Integer
  Dim vTmp As String
  
  If Rep.RepName <> "" Then
    Label12.Caption = "Filter Options for " & Rep.RepName
  End If
  Combo1.Clear
  Combo1.AddItem ""
  For vCnt = 1 To Rep.FieldCount
    vTmp = ""
    vTmp = Trim$(Rep.DataFields(vCnt).TableName) & "." & Trim$(Rep.DataFields(vCnt).FieldName) & LongString & Trim$(Rep.DataFields(vCnt).DateType)
    If Trim(vTmp) <> "" Then
      If Rep.DateFilter Then
        If InStr(1, UCase(vTmp), "DATE") > 0 Then
          Combo1.AddItem vTmp       'only add date fields to combo box
        End If
      Else
        Combo1.AddItem vTmp         'add all items
      End If
    End If
  Next vCnt
  
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

  Select Case vDataType
    Case 6, 3, 5                      'Numbers Only
      If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Then
      Else
        KeyAscii = 0
      End If
   
    Case 200, 201, 20, 202, 203       'String
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      
    Case Else
      KeyAscii = 0
  End Select

End Sub
