VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCloseDeal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Close Deal"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5160
   Icon            =   "frmCloseDeal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5040
      Top             =   1560
   End
   Begin VB.TextBox txtCD 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   2
      Tag             =   "Vehicle Make"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtCD 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   1
      Tag             =   "Vehicle Make"
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   525
      Left            =   120
      Picture         =   "frmCloseDeal.frx":1601A
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   525
      Left            =   3720
      Picture         =   "frmCloseDeal.frx":18CF5
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Save Records"
      Top             =   3120
      Width           =   1245
   End
   Begin VB.TextBox txtCD 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   5
      Tag             =   "Vehicle Make"
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtCD 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   3
      Tag             =   "Vehicle Make"
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtCD 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   0
      Tag             =   "Vehicle Make"
      Top             =   120
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   529
      _Version        =   393216
      Format          =   50987009
      CurrentDate     =   42460
   End
   Begin CarListing.ctlDone ctlDone1 
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   14
      Top             =   960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   -480
      X2              =   7680
      Y1              =   3060
      Y2              =   3060
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   2
      X1              =   -480
      X2              =   7800
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label6 
      Caption         =   "Buyer's ID Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Buyer's Last Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Sell Price:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Sell Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Buyer's Contact Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Buyer's First Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmCloseDeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

  Unload Me

End Sub

Private Sub Form_Load()

  txtCD(1).Text = CloseDeal.cldBuyerFirstName
  txtCD(2).Text = CloseDeal.cldBuyerLastName
  txtCD(3).Text = CloseDeal.cldBuyerID
  txtCD(4).Text = CloseDeal.cldBuyerContact
  If CloseDeal.cldBuyAmount <> 0 Then txtCD(5).Text = CloseDeal.cldBuyAmount
  DTPicker1.Value = Format(Now, "dd/mm/yyyy")
  Me.Caption = "Close Deal (" & CloseDeal.cldReg & ")"
  CheckID

End Sub

Private Sub Timer1_Timer()

  Timer1.Interval = 0
  If txtCD(1).Text = "" Then txtCD(1).SetFocus
  If txtCD(2).Text = "" Then txtCD(2).SetFocus
  If txtCD(3).Text = "" Then txtCD(3).SetFocus
  If txtCD(4).Text = "" Then txtCD(4).SetFocus
  If txtCD(5).Text = "" Then txtCD(5).SetFocus

End Sub

Private Sub txtCD_Change(Index As Integer)

  CheckID

End Sub

Private Sub txtCD_KeyPress(Index As Integer, KeyAscii As Integer)

  If Index = 5 Then
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or KeyAscii = 46 Then
      If InStr(1, txtCD(Index).Text, ".") > 0 And KeyAscii = 46 Then KeyAscii = 0
    Else
      KeyAscii = 0
    End If
  End If

End Sub

Private Sub cmdSave_Click()

  If txtCD(1).Text = "" Then MsgBox "Please complete the Buyer's First Name first. ", vbInformation + vbOKOnly, Me.Caption: Exit Sub
  If txtCD(2).Text = "" Then MsgBox "Please complete the Buyer's Last Name first. ", vbInformation + vbOKOnly, Me.Caption: Exit Sub
  If txtCD(3).Text = "" Then MsgBox "Please complete the Buyer's ID Number first. ", vbInformation + vbOKOnly, Me.Caption: Exit Sub
  If txtCD(4).Text = "" Then MsgBox "Please complete the Buyer's Contact Number first. ", vbInformation + vbOKOnly, Me.Caption: Exit Sub
  If Val(txtCD(5).Text) = 0 Then MsgBox "Please enter a Selling Price first. ", vbInformation + vbOKOnly, Me.Caption: Exit Sub
  CloseDeal.cldBuyerFirstName = txtCD(1).Text
  CloseDeal.cldBuyerLastName = txtCD(2).Text
  CloseDeal.cldBuyerID = txtCD(3).Text
  CloseDeal.cldBuyerContact = txtCD(4).Text
  CloseDeal.cldBuyAmount = Val(txtCD(5).Text)
  CloseDeal.cldBuyDate = GetDateVal(DTPicker1)
  CloseDeal.cldDone = True
  
  'close the deal and prevent saving
  AddClient.cmdSave.Enabled = False
  AddClient.cmdCloseDeal.Enabled = False
  Unload Me
  
  

End Sub

Private Sub CheckID()

  If ValidID(txtCD(3).Text) Then
    ctlDone1(0).Done = True
  Else
    ctlDone1(0).Done = False
  End If

End Sub
