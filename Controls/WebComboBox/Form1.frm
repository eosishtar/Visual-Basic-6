VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin WebComboBox.ctlWebCombo ctlWebCombo1 
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Top             =   1080
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   556
      vComboType      =   2
      vSearchInCombo  =   -1  'True
      vEnabled        =   -1  'True
      sSearchColor    =   12648447
      bAutoComplete   =   -1  'True
      vSearchText     =   "Search me..."
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Search Auto Complete"
      Height          =   495
      Left            =   7680
      TabIndex        =   6
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load Data"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Enabled"
      Height          =   495
      Left            =   7680
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Visible"
      Height          =   495
      Left            =   7680
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Search On / Off"
      Height          =   495
      Left            =   7680
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Combo Style"
      Height          =   495
      Left            =   7680
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim DatabasePath As String


Private Sub Check1_Click()
  If Check1.value <> 0 Then
    ctlWebCombo1.STYLE_COMBO = tDropDownCombo
  Else
    ctlWebCombo1.STYLE_COMBO = tDropdownList
  End If
End Sub

Private Sub Check2_Click()
  ctlWebCombo1.SEARCH = Check2.value
End Sub

Private Sub Check3_Click()
  ctlWebCombo1.Visible = Check3.value
End Sub

Private Sub Check4_Click()
  ctlWebCombo1.ENABLED_COMBO = Check4.value
End Sub

Private Sub Check5_Click()
  ctlWebCombo1.SEARCH_AUTOCOMPLETE = Check5.value
End Sub


Private Sub Command3_Click()

  ctlWebCombo1.ClearList
  Screen.MousePointer = vbHourglass
    
  'ctlWebCombo1.PopulateListSQL cn, "Category", "Description", False
  ctlWebCombo1.PopulateListSQL cn, "PCodes", "Field1", False
  
  ctlWebCombo1.AddItem "Grey"
  ctlWebCombo1.AddItem "Gray"
  ctlWebCombo1.AddItem "Grundge"
  ctlWebCombo1.AddItem "Yellow"

  Screen.MousePointer = vbNormal

End Sub

Private Sub ctlWebCombo1_ItemSelected(ItemID As Long, ItemName As String)
  MsgBox "ItemID " & "'" & ItemID & "'" & " - Item Name " & "'" & ItemName & "'"
End Sub

Private Sub Form_Load()

  Set cn = New ADODB.Connection
  
  DatabasePath = App.Path & "\DB.mdb"
  With cn
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .Open "Provider=microsoft.jet.OLEDB.4.0; ;" & "Data Source = " & DatabasePath & ";Jet OLEDB"
  End With


  ctlWebCombo1.STYLE_COMBO = tDropdownList

  
End Sub


Private Sub Form_Unload(Cancel As Integer)

  Set cn = Nothing

End Sub

