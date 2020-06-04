VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check6 
      Caption         =   "Force UCase"
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin TextBox.ctlTextbox ctlTextbox1 
      Height          =   300
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Test"
      Enabled         =   -1  'True
      MaxLength       =   0
      mTextBoxMode    =   1
      mForceUcase     =   0   'False
      TextAlignment   =   0
      Appearance      =   1
      BorderStyle     =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      PasswordChar    =   ""
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

Private Sub Command1_Click()
MsgBox "License Key Count " & LicenseKeyControl1.LicenseKeyCount
End Sub



Private Sub Check6_Click()
ctlTextbox1.ForceUCase = Check6.value
End Sub

Private Sub Combo1_Click()
ctlTextbox1.TextBoxMode = Combo1.ListIndex

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
  
  
  Combo1.Clear
  Combo1.AddItem "tbAllCharacters"
  Combo1.AddItem "tbOnlyNumbers"
  Combo1.AddItem "tbOnlyCharacter"
  Combo1.ListIndex = 0
  
 
  
End Sub


Private Sub Form_Unload(Cancel As Integer)

  Set cn = Nothing

End Sub

